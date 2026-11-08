import { NextRequest, NextResponse } from 'next/server'
import * as XLSX from 'xlsx'

export const runtime = 'nodejs'
export const maxDuration = 30

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

type Rows = unknown[][]
type ColMap = Record<string, number>

// メインファイル（請求明細表・納品書）の列名候補
const MAIN_COLS = {
  SALES_DATE: ['売上日'],
  SLIP_NO: ['伝票番号'],
  PRODUCT_CD: ['商品CD', '商品ｺｰﾄﾞ', '商品コード'],
  PRODUCT_NAME: ['商品名'],
  SPEC: ['規格'],
  QTY: ['数量'],
  UNIT: ['単位'],
  UNIT_PRICE: ['単価'],
  AMOUNT: ['金額'],
  TAX: ['税'],
}

// 単位補完用ファイル（受注問合せ・売上問合せ）の列名候補
const LOOKUP_COLS = {
  ORDER_NO: ['受注番号'],
  PRODUCT_CD: ['商品ｺｰﾄﾞ', '商品CD', '商品コード'],
  SPEC: ['規格'],
  UNIT: ['単位'],
}

const str = (v: unknown) => String(v ?? '').trim()

// 先頭10行からヘッダー行を探し、列名→列番号のマップを作る
function findHeader(rows: Rows, defs: Record<string, string[]>, required: string[]): { idx: number; cols: ColMap } | null {
  for (let i = 0; i < Math.min(rows.length, 10); i++) {
    const cells = (rows[i] ?? []).map(str)
    const cols: ColMap = {}
    for (const [key, names] of Object.entries(defs)) {
      const c = cells.findIndex((h) => names.includes(h))
      if (c >= 0) cols[key] = c
    }
    if (required.every((k) => k in cols)) return { idx: i, cols }
  }
  return null
}

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

const keyFull = (slip: string, cd: string, spec: string) => `${slip}|${cd}|${spec}`
const keySlipCd = (slip: string, cd: string) => `${slip}|${cd}`

class InputError extends Error {}

async function readRows(file: File): Promise<Rows> {
  try {
    const wb = XLSX.read(Buffer.from(await file.arrayBuffer()), { type: 'buffer', cellDates: false })
    return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' })
  } catch {
    throw new InputError(`「${file.name}」の読み込みに失敗しました`)
  }
}

// 枠（フォーム項目）ごとのファイル: main=請求明細表（必須）, juchu=受注問合せ, uriage=売上問合せ（任意）
const LOOKUP_SLOTS = [
  { field: 'juchu', label: '受注問合せ' },
  { field: 'uriage', label: '売上問合せ' },
]

export async function POST(req: NextRequest) {
  try {
    return await convert(await req.formData())
  } catch (e) {
    if (e instanceof InputError) return NextResponse.json({ error: e.message }, { status: 400 })
    throw e
  }
}

async function convert(formData: FormData) {
  const mainFile = formData.get('main')
  if (!(mainFile instanceof File)) {
    throw new InputError('請求明細表が選択されていません')
  }

  // 請求明細表（従来の納品書も可）
  const rows = await readRows(mainFile)
  const header = findHeader(rows, MAIN_COLS, ['SLIP_NO', 'PRODUCT_CD', 'PRODUCT_NAME', 'QTY', 'UNIT_PRICE'])
  if (!header) {
    throw new InputError(
      `「${mainFile.name}」は請求明細表の形式ではありません（伝票番号・商品CD・商品名・数量・単価の見出しが見つかりません）`,
    )
  }
  const main = { name: mainFile.name, rows, ...header }

  // 単位の対応表（伝票番号+商品CD+規格 を優先し、伝票番号+商品CD で補う）
  const unitByFull = new Map<string, string>()
  const unitBySlipCd = new Map<string, string>()
  const lookupNames: string[] = []

  for (const slot of LOOKUP_SLOTS) {
    const file = formData.get(slot.field)
    if (!(file instanceof File)) continue
    const lrows = await readRows(file)
    const lookup = findHeader(lrows, LOOKUP_COLS, ['ORDER_NO', 'PRODUCT_CD', 'UNIT'])
    if (!lookup) {
      throw new InputError(
        `「${file.name}」は${slot.label}の形式ではありません（受注番号・商品ｺｰﾄﾞ・単位の見出しが見つかりません）`,
      )
    }
    const { cols } = lookup
    for (const r of lrows.slice(lookup.idx + 1)) {
      const unit = str(r[cols.UNIT])
      if (!unit) continue
      const slip = str(r[cols.ORDER_NO])
      const cd = str(r[cols.PRODUCT_CD])
      const spec = 'SPEC' in cols ? str(r[cols.SPEC]) : ''
      if (!unitByFull.has(keyFull(slip, cd, spec))) unitByFull.set(keyFull(slip, cd, spec), unit)
      if (!unitBySlipCd.has(keySlipCd(slip, cd))) unitBySlipCd.set(keySlipCd(slip, cd), unit)
    }
    lookupNames.push(file.name)
  }

  const { cols } = main
  const get = (r: unknown[], key: string) => (key in cols ? r[cols[key]] : '')

  // A1の文字列に「アットライフ」が含まれる場合は仕入先コードを975にする
  const a1Value = str(rows[0]?.[0])
  const supplierCode = a1Value.includes('アットライフ') ? SUPPLIER_CODE_ATLIFE : SUPPLIER_CODE_DEFAULT

  const outputData: unknown[][] = [OUTPUT_HEADERS]
  const slipCounter: Record<string, number> = {}
  let missingUnit = 0

  for (const r of rows.slice(main.idx + 1)) {
    const productName = str(get(r, 'PRODUCT_NAME'))

    const productCd = str(get(r, 'PRODUCT_CD'))

    // Skip empty rows / 送料サービス rows
    if (!productName) continue
    if (productName === '送料サービス') continue

    const slipNo = str(get(r, 'SLIP_NO'))
    slipCounter[slipNo] = (slipCounter[slipNo] ?? 0) + 1
    const detailLine = slipCounter[slipNo]

    const taxStr = str(get(r, 'TAX'))
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

    const salesDate = toYYMMDD(get(r, 'SALES_DATE'))
    const spec = get(r, 'SPEC')
    const qty = Number(get(r, 'QTY'))
    const unitPrice = Number(get(r, 'UNIT_PRICE'))
    const tradeType = qty < 0
      ? TRADE_TYPE_RETURN
      : unitPrice < 0
        ? TRADE_TYPE_PRICE_ADJUST
        : TRADE_TYPE

    // 単位: 送料・返品手数料は「個」、それ以外はメインファイルの単位 → 受注/売上問合せから補完
    let unit: string
    if (productCd === '0020000000000' || productName === '返品手数料') {
      unit = '個'
    } else {
      unit = str(get(r, 'UNIT'))
        || unitByFull.get(keyFull(slipNo, productCd, str(spec)))
        || unitBySlipCd.get(keySlipCd(slipNo, productCd))
        || ''
      if (!unit) missingUnit++
    }

    outputData.push([
      supplierCode,            // A: 仕入先コード
      salesDate,               // B: 納品日
      slipNo,                  // C: 納品伝票No
      detailLine,              // D: 納品明細行
      tradeType,               // E: 仕入取引区分
      tradeType === TRADE_TYPE_PRICE_ADJUST
        ? PRICE_ADJUST_PRODUCT_CD
        : get(r, 'PRODUCT_CD'), // F: 仕入先商品コード
      productName,             // G: 商品名
      spec,                    // H: 仕様
      '',                      // I: カラー名
      '',                      // J: サイズ名
      '',                      // K: メーカー品番
      DEST_CODE,               // L: 納品先コード
      taxCategory,             // M: 税区分
      taxRate,                 // N: 税率
      get(r, 'QTY'),           // O: 数量
      unit,                    // P: 単位
      get(r, 'UNIT_PRICE'),    // Q: 単価
      get(r, 'AMOUNT'),        // R: 金額
      '',                      // S: 売上単価
      reducedTax,              // T: 軽減税率区分
      salesDate,               // U: 元の仕入計上日
    ])
  }

  if (outputData.length <= 1) {
    throw new InputError('変換対象のデータが見つかりませんでした')
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
      // 画面表示用の集計
      'X-Main-File': encodeURIComponent(main.name),
      'X-Row-Count': String(outputData.length - 1),
      'X-Lookup-Count': String(lookupNames.length),
      'X-Missing-Unit': String(missingUnit),
      'Access-Control-Expose-Headers': 'X-Main-File, X-Row-Count, X-Lookup-Count, X-Missing-Unit',
    },
  })
}
