import type { Metadata } from 'next'
import './globals.css'

export const metadata: Metadata = {
  title: 'TrackMyFund',
  description: 'Track your monthly SIP investments across mutual funds. Data stored in your own Excel file.',
}

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  )
}
