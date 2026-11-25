import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Outlook & Hotmail OAuth2 Mail Manager",
  description:
    "Hệ thống quản lý hộp thư Outlook/Hotmail qua Microsoft Graph API & IMAP XOAUTH2, bóc tách OTP tự động và Forward 24h theo lịch.",
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="vi" className="dark">
      <body className="bg-slate-950 text-slate-100 antialiased min-h-screen">
        {children}
      </body>
    </html>
  );
}
