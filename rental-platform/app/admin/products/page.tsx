'use client'
import { useState } from 'react'
import Link from 'next/link'
import { MOCK_PRODUCTS } from '@/lib/mock-data'
import { formatPrice, formatDate, getGradeLabel } from '@/lib/utils'
import { ArrowLeft, Plus, Search, Package } from 'lucide-react'
import type { ProductStatus } from '@/types'

const STATUS_OPTIONS: { value: ProductStatus | ''; label: string }[] = [
  { value: '', label: 'すべて' },
  { value: 'AVAILABLE', label: '在庫中' },
  { value: 'RENTED', label: '貸出中' },
  { value: 'MAINTENANCE', label: '整備中' },
  { value: 'RETIRED', label: '廃棄' },
]

const STATUS_STYLE: Record<ProductStatus, string> = {
  AVAILABLE: 'bg-green-100 text-green-700',
  RENTED: 'bg-blue-100 text-blue-700',
  MAINTENANCE: 'bg-yellow-100 text-yellow-700',
  RETIRED: 'bg-gray-100 text-gray-500',
}

const STATUS_LABEL: Record<ProductStatus, string> = {
  AVAILABLE: '在庫中',
  RENTED: '貸出中',
  MAINTENANCE: '整備中',
  RETIRED: '廃棄',
}

export default function AdminProductsPage() {
  const [search, setSearch] = useState('')
  const [statusFilter, setStatusFilter] = useState<ProductStatus | ''>('')
  const [showModal, setShowModal] = useState(false)

  const filtered = MOCK_PRODUCTS.filter(p => {
    if (statusFilter && p.status !== statusFilter) return false
    if (search && !p.name.includes(search) && !(p.brand ?? '').includes(search)) return false
    return true
  })

  return (
    <div className="max-w-6xl mx-auto px-4 py-10">
      <div className="flex items-center gap-4 mb-8">
        <Link href="/admin" className="text-gray-400 hover:text-gray-600">
          <ArrowLeft size={20} />
        </Link>
        <div className="flex-1">
          <h1 className="text-2xl font-bold text-[#1A5276]">商品管理</h1>
          <p className="text-gray-500 text-sm mt-1">在庫・整備状況の管理</p>
        </div>
        <button
          onClick={() => setShowModal(true)}
          className="flex items-center gap-2 bg-[#1A5276] hover:bg-blue-900 text-white text-sm font-bold px-4 py-2 rounded-lg transition-colors"
        >
          <Plus size={16} /> 商品登録
        </button>
      </div>

      {/* フィルター */}
      <div className="bg-white rounded-xl shadow-sm border border-gray-100 p-4 mb-6 flex flex-wrap gap-3">
        <div className="relative flex-1 min-w-48">
          <Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
          <input
            className="w-full pl-8 pr-3 py-2 border border-gray-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-[#1A5276]"
            placeholder="商品名・ブランドで検索"
            value={search}
            onChange={e => setSearch(e.target.value)}
          />
        </div>
        <div className="flex gap-2">
          {STATUS_OPTIONS.map(s => (
            <button
              key={s.value}
              onClick={() => setStatusFilter(s.value as ProductStatus | '')}
              className={`px-3 py-2 rounded-lg text-xs font-medium transition-colors ${statusFilter === s.value ? 'bg-[#1A5276] text-white' : 'bg-gray-100 text-gray-600 hover:bg-gray-200'}`}
            >
              {s.label}
            </button>
          ))}
        </div>
      </div>

      <p className="text-sm text-gray-500 mb-4">{filtered.length}件</p>

      {/* 商品テーブル */}
      <div className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead className="bg-gray-50">
              <tr>
                {['商品名', 'ブランド', 'カテゴリ', 'ランク', '月額', 'ステータス', '整備完了日', '回収場所', '操作'].map(h => (
                  <th key={h} className="px-4 py-3 text-left text-xs font-medium text-gray-500">{h}</th>
                ))}
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {filtered.map(p => (
                <tr key={p.id} className="hover:bg-gray-50">
                  <td className="px-4 py-3 font-medium text-gray-800">
                    <div className="flex items-center gap-2">
                      <span className="text-lg">
                        {{'冷蔵庫':'🧊','洗濯機':'🫧','電子レンジ':'📦'}[p.category] ?? '⚡'}
                      </span>
                      {p.name}
                    </div>
                  </td>
                  <td className="px-4 py-3 text-gray-500">{p.brand}</td>
                  <td className="px-4 py-3 text-gray-500">{p.category}</td>
                  <td className="px-4 py-3">
                    <span className={`text-xs font-bold px-2 py-0.5 rounded-full ${p.grade === 'A' ? 'bg-green-100 text-green-700' : 'bg-yellow-100 text-yellow-700'}`}>
                      {getGradeLabel(p.grade)}
                    </span>
                  </td>
                  <td className="px-4 py-3 text-gray-700 font-medium">{formatPrice(p.monthlyPrice)}</td>
                  <td className="px-4 py-3">
                    <span className={`text-xs font-bold px-2 py-0.5 rounded-full ${STATUS_STYLE[p.status]}`}>
                      {STATUS_LABEL[p.status]}
                    </span>
                  </td>
                  <td className="px-4 py-3 text-gray-500 text-xs">
                    {p.refurbishedAt ? formatDate(p.refurbishedAt) : '—'}
                  </td>
                  <td className="px-4 py-3 text-gray-500 text-xs">{p.sourceLocation ?? '—'}</td>
                  <td className="px-4 py-3">
                    <div className="flex gap-2">
                      <button className="text-xs text-[#1A5276] hover:underline">編集</button>
                      <button className="text-xs text-gray-400 hover:underline">ステータス変更</button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {/* 商品登録モーダル */}
      {showModal && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 px-4">
          <div className="bg-white rounded-2xl shadow-xl w-full max-w-lg p-8">
            <h2 className="text-lg font-bold text-gray-800 mb-6 flex items-center gap-2">
              <Package size={20} className="text-[#1A5276]" /> 新規商品登録
            </h2>
            <div className="space-y-4">
              {[
                { key: 'name', label: '商品名 *', placeholder: '例：シャープ 冷蔵庫 SJ-W352K' },
                { key: 'brand', label: 'ブランド *', placeholder: '例：SHARP' },
                { key: 'monthlyPrice', label: '月額料金（円）*', placeholder: '例：2500' },
                { key: 'sourceLocation', label: '回収場所', placeholder: '例：東京都練馬区' },
                { key: 'serialNo', label: 'シリアルNo', placeholder: '例：SN-20240101-001' },
              ].map(f => (
                <div key={f.key}>
                  <label className="text-xs font-medium text-gray-600">{f.label}</label>
                  <input
                    className="w-full border border-gray-200 rounded-lg px-3 py-2 mt-1 text-sm focus:outline-none focus:ring-2 focus:ring-[#1A5276]"
                    placeholder={f.placeholder}
                  />
                </div>
              ))}
              <div>
                <label className="text-xs font-medium text-gray-600">カテゴリ *</label>
                <select className="w-full border border-gray-200 rounded-lg px-3 py-2 mt-1 text-sm focus:outline-none focus:ring-2 focus:ring-[#1A5276]">
                  {['冷蔵庫', '洗濯機', '電子レンジ', 'その他家電'].map(c => (
                    <option key={c}>{c}</option>
                  ))}
                </select>
              </div>
              <div>
                <label className="text-xs font-medium text-gray-600">品質ランク *</label>
                <select className="w-full border border-gray-200 rounded-lg px-3 py-2 mt-1 text-sm focus:outline-none focus:ring-2 focus:ring-[#1A5276]">
                  <option value="A">Aランク（90点以上・ほぼ新品同様）</option>
                  <option value="B">Bランク（70〜89点・機能問題なし）</option>
                </select>
              </div>
              <div>
                <label className="text-xs font-medium text-gray-600">商品説明</label>
                <textarea
                  rows={3}
                  className="w-full border border-gray-200 rounded-lg px-3 py-2 mt-1 text-sm focus:outline-none focus:ring-2 focus:ring-[#1A5276] resize-none"
                  placeholder="商品の特徴・状態を記入してください"
                />
              </div>
            </div>
            <div className="flex gap-3 mt-6">
              <button
                onClick={() => setShowModal(false)}
                className="flex-1 border border-gray-300 text-gray-600 font-medium py-3 rounded-xl text-sm hover:bg-gray-50 transition-colors"
              >
                キャンセル
              </button>
              <button className="flex-1 bg-[#1A5276] hover:bg-blue-900 text-white font-bold py-3 rounded-xl text-sm transition-colors">
                登録する
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}
