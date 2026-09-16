import type { Metadata } from 'next';
import { Geist, Noto_Sans_Hebrew, Source_Serif_4 } from 'next/font/google';
import './globals.css';

const geist = Geist({ variable: '--font-geist-sans', subsets: ['latin', 'cyrillic'] });
const sourceSerif = Source_Serif_4({ variable: '--font-source-serif', subsets: ['latin', 'cyrillic'] });
const notoHebrew = Noto_Sans_Hebrew({ variable: '--font-noto-hebrew', subsets: ['hebrew'] });

export const metadata: Metadata = {
  title: 'Мой видуй — подготовка к Йом-Кипуру',
  description: 'Личный помощник по 24 частям короткого видуя «Ашамну».',
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return (
    <html lang="ru">
      <body className={`${geist.variable} ${sourceSerif.variable} ${notoHebrew.variable} antialiased`}>
        {children}
      </body>
    </html>
  );
}
