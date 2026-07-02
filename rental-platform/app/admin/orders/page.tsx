'use client'
import { useState } from 'react'
import Link from 'next/link'
import { MOCK_ORDERS } from '@/lib/mock-data'
import { formatPrice, formatDate, getStatusLabel } from '@/lib/utils'
import { ArrowLeft, Search, AlertCircle, CheckCircle, Clock } from 'lucide-react'
import type { OrderStatus } from '@/types'

const STATUS_OPTIONS: { value: OrderStatus | ''; label: string }[] = [
  { value: '', label: 'すべて' },
  { value: 'ACTIVE', label: '契約中' },
  { value: 'PENDING_RETURN', label: '返却待ち' },
  { value: 'COMPLETED', label: '完了' },
  { value: 'CANCELLED', label: 'キャンセル' },
]

const STATUS_STYLE: Record<OrderStatus, string> = {
  ACTIVE: 'bg-green-100 text-green-700',
  PENDING_RETURN: 'bg-orange-100 text-orange-700',
  COMPLETED: 'bg-gray-100 text-gray-500',
  CANCELLED: 'bg-red-100 text-red-500',
}

const STATUS_ICON: Record<OrderStatus, React.ReactNode> = {
  ACTIVE: <CheckCircle size={12} />,
  PENDING_RETURN: <Clock size={12} />,
  COMPLETED: <CheckCircle size={12} />,
  CANCELLED: <AlertCircle size={12} />,
}

export default function AdminOrdersPage() {
  const [search, setSearch] = useState('')
  const [statusFilter, setStatusFilter] = useState<OrderStatus | ''>('')

  const filtered = MOCK_ORDERS.filter(o => {
    if (statusFilter && o.status !== statusFilter) return false
    if (search && !o.product?.name.includes(search) && !o.id.includes(search)) return false
    return true
  })

  const pendingReturn = MOCK_ORDERS.filter(o => o.status === 'PENDING_RETURN')

  return (
    <div className="max-w-6xl mx-auto px-4 py-10">
      <div className="flex items-center gap-4 mb-8">
        <Link href="/admin" className="text-gray-400 hover:text-gray-600">
          <ArrowLeft size={20} />
        </Link>
        <div>
          <h1 className="text-2xl font-bold text-[#1A5276]">注文管理</h1>
          <p className="text-gray-500 text-sm mt-1">レンタル契約・返却の管理</p>
        </div>
      </div>

      {/* 返却待ちアラート */}
      {pendingReturn.length > 0 && (
        <div className="bg-orange-50 border border-orange-300 rounded-xl p-4 mb-6 flex items-start gap-3">
          <AlertCircle size={18} className="text-orange-600 flex-shrink-0 mt-0.5" />
          <div>
            <p className="font-bold text-orange-800 text-sm">{pendingReturn.length}件の返却申請があります</p>
            <p className="text-xs text-orange-700 mt-1">集荷の手配をしてください。</p>
          </div>
        </div>
      )}

      {/* KPIミニ */}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 mb-6">
        {[
          { label: '契約中', value: MOCK_ORDERS.filter(o => o.status === 'ACTIVE').length, color: 'text-green-600' },
          { label: '返却待ち', value: MOCK_ORDERS.filter(o => o.status === 'PENDING_RETURN').length, color: 'text-orange-600' },
          { label: '完了', value: MOCK_ORDERS.filter(o => o.status === 'COMPLETED').length, color: 'text-gray-500' },
          { label: '月次売上', value: formatPrice(MOCK_ORDERS.filter(o => o.status === 'ACTIVE').reduce((s, o) => s + o.monthlyPrice, 0)), color: 'text-[#1A5276]' },
        ].map(k => (
          <div key={k.label} className="bg-white rounded-xl shadow-sm border border-gray-100 p-4">
            <p className="text-xs text-gray-400">{k.label}</p>
            <p className={`text-xl font-bold mt-1 ${k.color}`}>{k.value}</p>
          </div>
        ))}
      </div>

      {/* フィルター */}
      <div className="bg-white rounded-xl shadow-sm border border-gray-100 p-4 mb-6 flex flex-wrap gap-3">
        <div className="relative flex-1 min-w-48">
          <Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
          <input
            className="w-full pl-8 pr-3 py-2 border border-gray-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-[#1A5276]"
            placeholder="注文ID・商品名で検索"
            value={search}
            onChange={e => setSearch(e.target.value)}
          />
        </div>
        <div className="flex flex-wrap gap-2">
          {STATUS_OPTIONS.map(s => (
            <button
              key={s.value}
              onClick={() => setStatusFilter(s.value as OrderStatus | '')}
              className={`px-3 py-2 rounded-lg text-xs font-medium transition-colors ${statusFilter === s.value ? 'bg-[#1A5276] text-white' : 'bg-gray-100 text-gray-600 hover:bg-gray-200'}`}
            >
              {s.label}
            </button>
          ))}
        </div>
      </div>

      <p className="text-sm text-gray-500 mb-4">{filtered.length}件</p>

      {/* 注文テーブル */}
      <div className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead className="bg-gray-50">
              <tr>
                {['注文ID', '商品名', '月額', 'ステータス', '開始日', '最低期間終了', '配送先', '操作'].map(h => (
                  <th key={h} className="px-4 py-3 text-left text-xs font-medium text-gray-500">{h}</th>
                ))}
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {filtered.length === 0 ? (
                <tr>
                  <td colSpan={8} className="px-4 py-12 text-center text-gray-400 text-sm">
                    該当する注文はありません
                  </td>
                </tr>
              ) : (
                filtered.map(o => (
                  <tr key={o.id} className="hover:bg-gray-50">
                    <td className="px-4 py-3 text-xs text-gray-400 font-mono">{o.id.slice(0, 8)}</td>
                    <td className="px-4 py-3 font-medium text-gray-800">{o.product?.name ?? '—'}</td>
                    <td className="px-4 py-3 text-gray-700">{formatPrice(o.monthlyPrice)}</td>
                    <td className="px-4 py-3">
                      <span className={`inline-flex items-center gap-1 text-xs font-bold px-2 py-0.5 rounded-full ${STATUS_STYLE[o.status]}`}>
                        {STATUS_ICON[o.status]}{getStatusLabel(o.status)}
                      </span>
                    </td>
                    <td className="px-4 py-3 text-gray-500 text-xs">{formatDate(o.startDate)}</td>
                    <td className="px-4 py-3 text-gray-500 text-xs">{formatDate(o.minEndDate)}</td>
                    <td className="px-4 py-3 text-gray-500 text-xs">{o.deliveryAddress ?? '—'}</td>
                    <td className="px-4 py-3">
                      <div className="flex gap-2">
                        {o.status === 'PENDING_RETURN' && (
                          <button className="text-xs bg-orange-100 text-orange-700 hover:bg-orange-200 px-2 py-1 rounded font-medium transition-colors">
                            集荷手配
                          </button>
                        )}
                        {o.status === 'ACTIVE' && (
                          <button className="text-xs text-gray-400 hover:underline">詳細</button>
                        )}
                      </div>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  )
}
