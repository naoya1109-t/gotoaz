'use client'

import { useState, useRef } from 'react'

type SlotKey = 'main' | 'juchu' | 'uriage'

const SLOTS: { key: SlotKey; label: string; required: boolean; note: string }[] = [
  { key: 'main', label: '請求明細表', required: true, note: '変換元のデータ' },
  { key: 'juchu', label: '受注問合せ', required: false, note: '単位の補完用' },
  { key: 'uriage', label: '売上問合せ', required: false, note: '単位の補完用' },
]

const isExcel = (f: File) => /\.xlsx?$/i.test(f.name)

function FileSlot({
  label,
  required,
  note,
  file,
  onSelect,
  onClear,
}: {
  label: string
  required: boolean
  note: string
  file: File | null
  onSelect: (f: File) => void
  onClear: () => void
}) {
  const inputRef = useRef<HTMLInputElement>(null)
  const [dragOver, setDragOver] = useState(false)

  const pick = (list: FileList | null) => {
    const f = Array.from(list ?? []).find(isExcel)
    if (f) onSelect(f)
  }

  return (
    <div style={{ marginBottom: 16 }}>
      <div style={{ display: 'flex', alignItems: 'baseline', gap: 8, marginBottom: 6 }}>
        <span style={{ fontSize: 14, fontWeight: 600, color: '#333' }}>{label}</span>
        <span
          style={{
            fontSize: 11,
            padding: '1px 6px',
            borderRadius: 4,
            color: '#fff',
            background: required ? '#d32f2f' : '#9e9e9e',
          }}
        >
          {required ? '必須' : '任意'}
        </span>
        <span style={{ fontSize: 12, color: '#888' }}>{note}</span>
      </div>
      <div
        onClick={() => inputRef.current?.click()}
        onDragOver={(e) => {
          e.preventDefault()
          setDragOver(true)
        }}
        onDragLeave={() => setDragOver(false)}
        onDrop={(e) => {
          e.preventDefault()
          setDragOver(false)
          pick(e.dataTransfer.files)
        }}
        style={{
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          gap: 8,
          minHeight: 56,
          padding: '12px 16px',
          border: '2px dashed',
          borderColor: dragOver ? '#1976d2' : file ? '#4caf50' : '#ccc',
          borderRadius: 8,
          background: dragOver ? '#e3f2fd' : file ? '#f1f8e9' : '#fafafa',
          cursor: 'pointer',
          transition: 'all 0.2s',
        }}
      >
        <span style={{ fontSize: 14, color: file ? '#2e7d32' : '#888', wordBreak: 'break-all' }}>
          {file
            ? `${file.name}（${(file.size / 1024).toFixed(1)} KB）`
            : 'ドラッグ＆ドロップ、またはクリックして選択'}
        </span>
        {file && (
          <button
            onClick={(e) => {
              e.stopPropagation()
              onClear()
            }}
            aria-label={`${label}を取り消し`}
            style={{ border: 'none', background: 'transparent', color: '#888', cursor: 'pointer', fontSize: 14 }}
          >
            ✕
          </button>
        )}
        <input
          ref={inputRef}
          type="file"
          accept=".xls,.xlsx"
          style={{ display: 'none' }}
          onChange={(e) => {
            pick(e.target.files)
            e.target.value = ''
          }}
        />
      </div>
    </div>
  )
}

export default function Home() {
  const [files, setFiles] = useState<Record<SlotKey, File | null>>({ main: null, juchu: null, uriage: null })
  const [status, setStatus] = useState<'idle' | 'loading' | 'success' | 'error'>('idle')
  const [errorMsg, setErrorMsg] = useState('')
  const [resultMsg, setResultMsg] = useState('')

  const setSlot = (key: SlotKey, f: File | null) => {
    setFiles((prev) => ({ ...prev, [key]: f }))
    setStatus('idle')
    setErrorMsg('')
  }

  const handleConvert = async () => {
    if (!files.main) return
    setStatus('loading')
    setErrorMsg('')

    const formData = new FormData()
    for (const s of SLOTS) {
      const f = files[s.key]
      if (f) formData.append(s.key, f)
    }

    try {
      const res = await fetch('/api/convert', { method: 'POST', body: formData })
      if (!res.ok) {
        const json = await res.json().catch(() => ({}))
        throw new Error(json.error || `サーバーエラー: ${res.status}`)
      }
      const rowCount = res.headers.get('X-Row-Count') ?? '0'
      const lookupCount = Number(res.headers.get('X-Lookup-Count') ?? 0)
      const missing = Number(res.headers.get('X-Missing-Unit') ?? 0)

      const blob = await res.blob()
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = `実績表_${files.main.name.replace(/\.[^.]+$/, '')}.xlsx`
      a.click()
      URL.revokeObjectURL(url)

      let msg = `変換が完了しました（${rowCount}行）。ダウンロードが開始されます。`
      if (lookupCount === 0) msg += '\n※ 受注問合せ・売上問合せがないため、単位は空欄です。'
      else if (missing > 0) msg += `\n※ 単位が見つからなかった行が ${missing} 行あります（空欄で出力）。`
      setResultMsg(msg)
      setStatus('success')
    } catch (err) {
      setErrorMsg(err instanceof Error ? err.message : '変換中にエラーが発生しました')
      setStatus('error')
    }
  }

  const canConvert = !!files.main && status !== 'loading'

  const styles = {
    container: {
      maxWidth: 600,
      margin: '60px auto',
      padding: '0 20px',
    },
    card: {
      background: '#fff',
      borderRadius: 12,
      padding: 40,
      boxShadow: '0 2px 12px rgba(0,0,0,0.1)',
    },
    title: {
      fontSize: 24,
      fontWeight: 700,
      marginBottom: 8,
      color: '#1a1a2e',
    },
    subtitle: {
      fontSize: 14,
      color: '#666',
      marginBottom: 32,
    },
    button: {
      display: 'block',
      width: '100%',
      marginTop: 24,
      padding: '14px 0',
      fontSize: 16,
      fontWeight: 600,
      border: 'none',
      borderRadius: 8,
      cursor: canConvert ? 'pointer' : 'not-allowed',
      background: canConvert ? '#1976d2' : '#ccc',
      color: '#fff',
      transition: 'background 0.2s',
    },
    successMsg: {
      marginTop: 16,
      padding: '12px 16px',
      background: '#e8f5e9',
      border: '1px solid #a5d6a7',
      borderRadius: 6,
      color: '#2e7d32',
      fontSize: 14,
      whiteSpace: 'pre-line' as const,
    },
    errorMsg: {
      marginTop: 16,
      padding: '12px 16px',
      background: '#ffebee',
      border: '1px solid #ef9a9a',
      borderRadius: 6,
      color: '#c62828',
      fontSize: 14,
    },
    rules: {
      marginTop: 32,
      padding: '16px',
      background: '#f8f9fa',
      borderRadius: 8,
      fontSize: 13,
      color: '#555',
    },
    ruleTitle: {
      fontWeight: 600,
      marginBottom: 8,
      color: '#333',
    },
    ruleItem: {
      marginBottom: 4,
    },
  }

  return (
    <div style={styles.container}>
      <div style={styles.card}>
        <h1 style={styles.title}>納品書変換ツール</h1>
        <p style={styles.subtitle}>請求明細表Excelファイルをアップロードし、ゴトウ・アズ・プランニング様用にフォーマットに変換してダウンロードします</p>

        {SLOTS.map((s) => (
          <FileSlot
            key={s.key}
            label={s.label}
            required={s.required}
            note={s.note}
            file={files[s.key]}
            onSelect={(f) => setSlot(s.key, f)}
            onClear={() => setSlot(s.key, null)}
          />
        ))}

        <button
          style={styles.button}
          onClick={handleConvert}
          disabled={!canConvert}
        >
          {status === 'loading' ? '変換中...' : '変換してダウンロード'}
        </button>

        {status === 'success' && <div style={styles.successMsg}>{resultMsg}</div>}
        {status === 'error' && (
          <div style={styles.errorMsg}>
            エラー: {errorMsg}
          </div>
        )}

        <div style={styles.rules}>
          <p style={styles.ruleTitle}>変換ルール</p>
          <p style={styles.ruleItem}>・「商品名」が「送料サービス」の行は除外</p>
          <p style={styles.ruleItem}>・「商品名」→ 新ファイルG列「商品名」、「規格」→ H列「仕様」</p>
          <p style={styles.ruleItem}>・「税」→ 新ファイルT列「軽減税率区分」（課10 = 0、課8* = 1）</p>
          <p style={styles.ruleItem}>・「税」→ 新ファイルN列「税率」（課10 = 10、課8* = 8）</p>
          <p style={styles.ruleItem}>・A1に「アットライフ」を含む場合、新ファイルA列「仕入先コード」= 975（それ以外は976）</p>
          <p style={styles.ruleItem}>・「数量」がマイナスの場合、新ファイルE列「仕入取引区分」= 5</p>
          <p style={styles.ruleItem}>・「単価」がマイナスの場合、新ファイルE列「仕入取引区分」= 6、新ファイルF列「仕入先商品コード」= 0040000000000</p>
          <p style={styles.ruleItem}>・新ファイルP列「単位」は、受注問合せ・売上問合せの「受注番号＋商品コード（＋規格）」が請求明細表の「伝票番号＋商品CD（＋規格）」と一致する行から補完（一致しない行も単位を空欄にして出力）</p>
        </div>
      </div>
    </div>
  )
}
