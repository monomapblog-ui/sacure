'use client'
import Link from 'next/link'
import { useState } from 'react'
import { Menu, X, RefreshCw } from 'lucide-react'

export default function Navbar() {
  const [open, setOpen] = useState(false)

  return (
    <nav className="bg-[#1A5276] text-white shadow-lg">
      <div className="max-w-6xl mx-auto px-4 flex items-center justify-between h-16">
        <Link href="/" className="flex items-center gap-2 font-bold text-xl">
          <RefreshCw size={22} className="text-green-400" />
          <span>サキュレ</span>
          <span className="text-green-400 text-sm font-normal">レンタル</span>
        </Link>

        <div className="hidden md:flex items-center gap-6 text-sm">
          <Link href="/" className="hover:text-green-300 transition-colors">商品一覧</Link>
          <Link href="/mypage" className="hover:text-green-300 transition-colors">マイページ</Link>
          <Link href="/auth/login" className="bg-green-600 hover:bg-green-500 px-4 py-2 rounded-md transition-colors">
            ログイン
          </Link>
          <Link href="/auth/register" className="border border-white hover:bg-white hover:text-[#1A5276] px-4 py-2 rounded-md transition-colors">
            新規登録
          </Link>
        </div>

        <button className="md:hidden" onClick={() => setOpen(!open)}>
          {open ? <X size={24} /> : <Menu size={24} />}
        </button>
      </div>

      {open && (
        <div className="md:hidden bg-[#154360] px-4 py-4 flex flex-col gap-3 text-sm">
          <Link href="/" onClick={() => setOpen(false)}>商品一覧</Link>
          <Link href="/mypage" onClick={() => setOpen(false)}>マイページ</Link>
          <Link href="/auth/login" onClick={() => setOpen(false)}>ログイン</Link>
          <Link href="/auth/register" onClick={() => setOpen(false)}>新規登録</Link>
        </div>
      )}
    </nav>
  )
}
