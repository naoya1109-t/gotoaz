'use client'

import { useState, useRef } from 'react'

export default function Home() {
  const [file, setFile] = useState<File | null>(null)
  const [status, setStatus] = useState<'idle' | 'loading' | 'success' | 'error'>('idle')
  const [errorMsg, setErrorMsg] = useState('')
  const inputRef = useRef<HTMLInputElement>(null)

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0]
    if (f) {
      setFile(f)
      setStatus('idle')
      setErrorMsg('')
    }
  }

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault()
    const f = e.dataTransfer.files[0]
    if (f && (f.name.endsWith('.xls') || f.name.endsWith('.xlsx'))) {
      setFile(f)
      setStatus('idle')
      setErrorMsg('')
    }
  }

  const handleConvert = async () => {
    if (!file) return
    setStatus('loading')
    setErrorMsg('')

    const formData = new FormData()
    formData.append('file', file)

    try {
      const res = await fetch('/api/convert', { method: 'POST', body: formData })
      if (!res.ok) {
        const json = await res.json().catch(() => ({}))
        throw new Error(json.error || `サーバーエラー: ${res.status}`)
      }
      const blob = await res.blob()
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      const originalName = file.name.replace(/\.[^.]+$/, '')
      a.download = `実績表_${originalName}.xlsx`
      a.click()
      URL.revokeObjectURL(url)
      setStatus('success')
    } catch (err) {
      setErrorMsg(err instanceof Error ? err.message : '変換中にエラーが発生しました')
      setStatus('error')
    }
  }

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
    dropzone: {
      border: '2px dashed',
      borderColor: file ? '#4caf50' : '#ccc',
      borderRadius: 8,
      padding: '40px 20px',
      textAlign: 'center' as const,
      cursor: 'pointer',
      background: file ? '#f1f8e9' : '#fafafa',
      transition: 'all 0.2s',
    },
    dropzoneText: {
      fontSize: 15,
      color: file ? '#2e7d32' : '#888',
      marginBottom: 8,
    },
    fileInfo: {
      fontSize: 13,
      color: '#555',
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
      cursor: file && status !== 'loading' ? 'pointer' : 'not-allowed',
      background: file && status !== 'loading' ? '#1976d2' : '#ccc',
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
        <p style={styles.subtitle}>納品書ExcelファイルをアップロードAし、実績表フォーマットに変換してダウンロードします</p>

        <div
          style={styles.dropzone}
          onClick={() => inputRef.current?.click()}
          onDragOver={(e) => e.preventDefault()}
          onDrop={handleDrop}
        >
          <p style={styles.dropzoneText}>
            {file ? `選択済み: ${file.name}` : 'ここにファイルをドラッグ＆ドロップ、またはクリックして選択'}
          </p>
          {file && (
            <p style={styles.fileInfo}>
              サイズ: {(file.size / 1024).toFixed(1)} KB
            </p>
          )}
          <input
            ref={inputRef}
            type="file"
            accept=".xls,.xlsx"
            style={{ display: 'none' }}
            onChange={handleFileChange}
          />
        </div>

        <button
          style={styles.button}
          onClick={handleConvert}
          disabled={!file || status === 'loading'}
        >
          {status === 'loading' ? '変換中...' : '変換してダウンロード'}
        </button>

        {status === 'success' && (
          <div style={styles.successMsg}>
            変換が完了しました。ダウンロードが開始されます。
          </div>
        )}
        {status === 'error' && (
          <div style={styles.errorMsg}>
            エラー: {errorMsg}
          </div>
        )}

        <div style={styles.rules}>
          <p style={styles.ruleTitle}>変換ルール</p>
          <p style={styles.ruleItem}>・H列「商品名」が「送料サービス」の行は除外</p>
          <p style={styles.ruleItem}>・H列「商品名」→ 新ファイルG列「商品名」</p>
          <p style={styles.ruleItem}>・N列「税」→ 新ファイルT列「軽減税率区分」（課10 = 0、課8* = 1）</p>
          <p style={styles.ruleItem}>・N列「税」→ 新ファイルN列「税率」（課10 = 10、課8* = 8）</p>
        </div>
      </div>
    </div>
  )
}
