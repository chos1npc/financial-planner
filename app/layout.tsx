import type { Metadata } from 'next';
import './globals.css';

const rawSiteUrl = process.env.NEXT_PUBLIC_SITE_URL ?? 'http://localhost:3000/';
const siteUrl = rawSiteUrl.endsWith('/') ? rawSiteUrl : `${rawSiteUrl}/`;

export const metadata: Metadata = {
  metadataBase: new URL(siteUrl),
  title: '忙季算盤｜財報人力與完工預估',
  description: '用上一季財報量安排人力，並依每日實績即時估算忙季是否能如期完成。',
  openGraph: {
    title: '忙季算盤｜這個忙季，做得完嗎？',
    description: '換算歷史 T 日財報量、安排人力，並用每日實績即時估算是否能如期完成。',
    type: 'website',
    locale: 'zh_TW',
    images: [{ url: './og.png', width: 1200, height: 630, alt: '忙季算盤' }],
  },
  twitter: {
    card: 'summary_large_image',
    title: '忙季算盤｜這個忙季，做得完嗎？',
    description: '換算歷史 T 日財報量、安排人力，並用每日實績即時估算是否能如期完成。',
    images: ['./og.png'],
  },
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="zh-Hant">
      <body>{children}</body>
    </html>
  );
}
