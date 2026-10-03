import type { Metadata } from 'next';
import './globals.css';

export const metadata: Metadata = {
  title: 'Excel AI Agent',
  description: 'AI-powered Excel automation with evidence-based cell targeting.',
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
