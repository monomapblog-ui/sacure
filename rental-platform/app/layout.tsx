import type { Metadata } from "next";
import { Geist } from "next/font/google";
import "./globals.css";
import Navbar from "@/components/Navbar";

const geist = Geist({ subsets: ["latin"] });

export const metadata: Metadata = {
  title: "サキュレレンタル | 家電の月額レンタルサービス",
  description: "回収・再生した家電を月額でレンタル。冷蔵庫・洗濯機・電子レンジほか。株式会社サキュレの循環型家電レンタルサービス。",
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="ja">
      <body className={`${geist.className} bg-gray-50 min-h-screen flex flex-col`}>
        <Navbar />
        <main className="flex-1">{children}</main>
        <footer className="bg-[#1A5276] text-white mt-16 py-8 text-center text-sm">
          <p className="font-bold mb-1">株式会社サキュレ</p>
          <p className="text-gray-300">東京都大田区仲池上一丁目24番7号 | www.390.co.jp</p>
          <p className="text-gray-400 mt-2 text-xs">© 2026 Sakure Inc. All rights reserved.</p>
        </footer>
      </body>
    </html>
  );
}
