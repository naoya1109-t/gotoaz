import { NextRequest, NextResponse } from 'next/server'
import * as XLSX from 'xlsx'

export const runtime = 'nodejs'
export const maxDuration = 30

// Input column indices (0-based): A=0, B=1, ..., H=7, N=13
const IN_COL = {
  JUCHU_DATE: 0,   // A: 受注日
  SALES_DATE: 1,   // B: 売上日
  SLIP_NO: 2,      // C: 伝票番号
  CUSTOMER_ORDER: 3, // D: 客先注文番号
  CUSTOMER_NAME: 4,  // E: 得意先名
  DEST_NAME: 5,    // F: 送付先名
  PRODUCT_CD: 6,   // G: 商品CD
  PRODUCT_NAME: 7, // H: 商品名
  SPEC: 8,         // I: 規格
  QTY: 9,          // J: 数量
  UNIT: 10,        // K: 単位
  UNIT_PRICE: 11,  // L: 単価
  AMOUNT: 12,      // M: 金額
  TAX: 13,         // N: 税
  REMARKS: 14,     // O: 行備考
  RETAIL_PRICE: 15, // P: 小売価格
  WEB_CD: 16,      // Q: WEBCD
}

// Output column headers (A-U = 0-20)
const OUTPUT_HEADERS = [
  '仕入先コード',   // A
  '納品日',         // B
  '納品伝票No',     // C
  '納品明細行',     // D
  '仕入取引区分',   // E
  '仕入先商品コード', // F
  '商品名',         // G
  '仕様',           // H
  'カラー名',       // I
  'サイズ名',       // J
  'メーカー品番',   // K
  '納品先コード',   // L
  '税区分',         // M
  '税率',           // N
  '数量',           // O
  '単位',           // P
  '単価',           // Q
  '金額',           // R
  '売上単価',       // S
  '軽減税率区分',   // T
  '元の仕入計上日', // U
]

const SUPPLIER_CODE_DEFAULT = 976
const SUPPLIER_CODE_ATLIFE = 975
const DEST_CODE = 317876
const TAX_CATEGORY = '外税'
const TRADE_TYPE = 1
const TRADE_TYPE_RETURN = 5
const TRADE_TYPE_PRICE_ADJUST = 6
const PRICE_ADJUST_PRODUCT_CD = '0040000000000'

// Excelシリアル日付値を yy/mm/dd 文字列に変換
function toYYMMDD(val: unknown): string {
  if (typeof val === 'number' && val > 0) {
    const date = new Date(Math.round((val - 25569) * 86400000))
    const yy = String(date.getUTCFullYear()).slice(-2)
    const mm = String(date.getUTCMonth() + 1).padStart(2, '0')
    const dd = String(date.getUTCDate()).padStart(2, '0')
    return `${yy}/${mm}/${dd}`
  }
  return String(val ?? '')
}

export async function POST(req: NextRequest) {
  const formData = await req.formData()
  const file = formData.get('file') as File | null
  if (!file) {
    return NextResponse.json({ error: 'ファイルが選択されていません' }, { status: 400 })
  }

  const arrayBuffer = await file.arrayBuffer()
  const buffer = Buffer.from(arrayBuffer)

  let wb: XLSX.WorkBook
  try {
    wb = XLSX.read(buffer, { type: 'buffer', cellDates: false })
  } catch {
    return NextResponse.json({ error: 'Excelファイルの読み込みに失敗しました' }, { status: 400 })
  }

  const sheetName = wb.SheetNames[0]
  const ws = wb.Sheets[sheetName]
  const rows: unknown[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' })

  // A1の文字列に「アットライフ」が含まれる場合は仕入先コードを975にする
  const a1Value = String(rows[0]?.[0] ?? '')
  const supplierCode = a1Value.includes('アットライフ') ? SUPPLIER_CODE_ATLIFE : SUPPLIER_CODE_DEFAULT

  // Find header row (look for row containing 商品名 in H column)
  let headerRowIdx = -1
  for (let i = 0; i < Math.min(rows.length, 10); i++) {
    const row = rows[i] as unknown[]
    if (String(row[IN_COL.PRODUCT_NAME] ?? '').trim() === '商品名') {
      headerRowIdx = i
      break
    }
  }
  if (headerRowIdx === -1) {
    return NextResponse.json({ error: 'ヘッダー行（商品名）が見つかりません' }, { status: 400 })
  }

  const dataRows = rows.slice(headerRowIdx + 1)
  const outputData: unknown[][] = [OUTPUT_HEADERS]

  // Track sequential counter per slip number
  const slipCounter: Record<string, number> = {}

  for (const row of dataRows) {
    const r = row as unknown[]
    const productName = String(r[IN_COL.PRODUCT_NAME] ?? '').trim()

    // Skip empty rows
    if (!productName) continue

    // Skip 送料サービス rows
    if (productName === '送料サービス') continue

    const slipNo = String(r[IN_COL.SLIP_NO] ?? '').trim()
    slipCounter[slipNo] = (slipCounter[slipNo] ?? 0) + 1
    const detailLine = slipCounter[slipNo]

    const taxStr = String(r[IN_COL.TAX] ?? '').trim()
    // 課10 → taxRate=10, reduced=0; 課8* → taxRate=8, reduced=1; 非 → 非課税, taxRate=10
    let taxRate: number
    let reducedTax: number
    let taxCategory: string
    if (taxStr === '非') {
      taxRate = 10
      reducedTax = 0
      taxCategory = '非課税'
    } else if (taxStr === '課10') {
      taxRate = 10
      reducedTax = 0
      taxCategory = TAX_CATEGORY
    } else if (taxStr.startsWith('課8')) {
      taxRate = 8
      reducedTax = 1
      taxCategory = TAX_CATEGORY
    } else {
      taxRate = 0
      reducedTax = 0
      taxCategory = TAX_CATEGORY
    }

    const salesDate = toYYMMDD(r[IN_COL.SALES_DATE])
    const qty = Number(r[IN_COL.QTY])
    const unitPrice = Number(r[IN_COL.UNIT_PRICE])
    const tradeType = qty < 0
      ? TRADE_TYPE_RETURN
      : unitPrice < 0
        ? TRADE_TYPE_PRICE_ADJUST
        : TRADE_TYPE

    const outRow = [
      supplierCode,            // A: 仕入先コード
      salesDate,               // B: 納品日
      slipNo,                  // C: 納品伝票No
      detailLine,              // D: 納品明細行
      tradeType,               // E: 仕入取引区分
      tradeType === TRADE_TYPE_PRICE_ADJUST
        ? PRICE_ADJUST_PRODUCT_CD
        : r[IN_COL.PRODUCT_CD], // F: 仕入先商品コード
      productName,             // G: 商品名
      r[IN_COL.SPEC],         // H: 仕様
      '',                      // I: カラー名
      '',                      // J: サイズ名
      '',                      // K: メーカー品番
      DEST_CODE,               // L: 納品先コード
      taxCategory,             // M: 税区分
      taxRate,                 // N: 税率
      r[IN_COL.QTY],          // O: 数量
      (String(r[IN_COL.PRODUCT_CD] ?? '').trim() === '0020000000000' || productName === '返品手数料')
        ? '個'
        : r[IN_COL.UNIT],    // P: 単位
      r[IN_COL.UNIT_PRICE],   // Q: 単価
      r[IN_COL.AMOUNT],       // R: 金額
      '',                      // S: 売上単価
      reducedTax,              // T: 軽減税率区分
      salesDate,               // U: 元の仕入計上日
    ]

    outputData.push(outRow)
  }

  if (outputData.length <= 1) {
    return NextResponse.json({ error: '変換対象のデータが見つかりませんでした' }, { status: 400 })
  }

  const outWs = XLSX.utils.aoa_to_sheet(outputData)
  const outWb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(outWb, outWs, '実績表')

  const outBuffer = XLSX.write(outWb, { type: 'buffer', bookType: 'xlsx' })

  return new NextResponse(outBuffer, {
    status: 200,
    headers: {
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': 'attachment; filename="output.xlsx"',
    },
  })
}
