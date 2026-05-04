import type { Metadata } from 'next'

export const metadata: Metadata = {
  title: '納品書変換ツール',
  description: '納品書Excelファイルを実績表フォーマットに変換します',
}

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="ja">
      <body style={{ margin: 0, fontFamily: 'sans-serif', background: '#f5f5f5' }}>
        {children}
      </body>
    </html>
  )
}
