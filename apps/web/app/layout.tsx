import './globals.css';
import type { Metadata } from 'next';

export const metadata: Metadata = {
  title: 'Lokalmart Studio2',
  description: 'Import Export Odoo XLSX Studio'
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="id">
      <body>{children}</body>
    </html>
  );
}
