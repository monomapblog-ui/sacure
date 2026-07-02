'use client'
import { MOCK_ORDERS } from '@/lib/mock-data'
import { formatPrice, formatDate, getStatusLabel } from '@/lib/utils'
import { Package, Calendar, AlertCircle, ArrowRight } from 'lucide-react'
import Link from 'next/link'

export default function MyPage() {
  const orders = MOCK_ORDERS
  const activeOrders = orders.filter(o => o.status === 'ACTIVE')

  return (
    <div className="max-w-4xl mx-auto px-4 py-10">
      <h1 className="text-2xl font-bold text-[#1A5276] mb-2">マイページ</h1>
      <p className="text-gray-500 text-sm mb-8">ゲストプレビュー（ログインで全機能利用可）</p>

      {/* KYC ステータス */}
      <div className="bg-yellow-50 border border-yellow-300 rounded-xl p-4 mb-6 flex items-start gap-3">
        <AlertCircle size={20} className="text-yellow-600 flex-shrink-0 mt-0.5" />
        <div>
          <p className="font-bold text-yellow-800 text-sm">本人確認が未完了です</p>
          <p className="text-xs text-yellow-700 mt-1">運転免許証等をアップロードして本人確認を完了してください。確認完了後にレンタル申込が可能になります。</p>
          <button className="mt-2 text-xs bg-yellow-500 hover:bg-yellow-400 text-white px-3 py-1.5 rounded-lg font-medium transition-colors">
            本人確認書類をアップロード
          </button>
        </div>
      </div>

      {/* 契約中の商品 */}
      <section className="mb-8">
        <h2 className="text-lg font-bold text-gray-700 mb-4 flex items-center gap-2">
          <Package size={20} className="text-[#1A5276]" /> 現在のご契約
        </h2>
        {activeOrders.length === 0 ? (
          <div className="bg-white rounded-xl border border-dashed border-gray-300 p-10 text-center text-gray-400">
            <p className="text-4xl mb-3">📦</p>
            <p className="font-medium">現在レンタル中の商品はありません</p>
            <Link href="/" className="mt-4 inline-block text-sm text-[#1A5276] underline">商品を探す</Link>
          </div>
        ) : (
          <div className="space-y-4">
            {activeOrders.map(order => (
              <div key={order.id} className="bg-white rounded-xl shadow-sm border border-gray-100 p-5">
                <div className="flex items-start justify-between">
                  <div>
                    <span className="text-xs bg-green-100 text-green-700 font-bold px-2 py-0.5 rounded-full">
                      {getStatusLabel(order.status)}
                    </span>
                    <h3 className="font-bold text-gray-800 mt-2">{order.product?.name}</h3>
                    <p className="text-xs text-gray-500 mt-1">{order.product?.brand} / {order.product?.category}</p>
                  </div>
                  <p className="text-xl font-bold text-[#1A5276]">{formatPrice(order.monthlyPrice)}<span className="text-xs font-normal text-gray-400">/月</span></p>
                </div>
                <div className="mt-4 pt-4 border-t border-gray-100 grid grid-cols-2 gap-3 text-sm">
                  <div className="flex items-center gap-2 text-gray-600">
                    <Calendar size={14} />
                    <span>開始：{formatDate(order.startDate)}</span>
                  </div>
                  <div className="flex items-center gap-2 text-gray-600">
                    <Calendar size={14} />
                    <span>最低期間終了：{formatDate(order.minEndDate)}</span>
                  </div>
                </div>
                <div className="mt-4 flex gap-3">
                  <button className="flex-1 border border-red-300 text-red-600 hover:bg-red-50 text-sm font-medium py-2 rounded-lg transition-colors">
                    返却申請
                  </button>
                  <button className="flex-1 border border-gray-300 text-gray-600 hover:bg-gray-50 text-sm font-medium py-2 rounded-lg transition-colors">
                    トラブル報告
                  </button>
                </div>
              </div>
            ))}
          </div>
        )}
      </section>

      {/* 過去の履歴 */}
      <section>
        <h2 className="text-lg font-bold text-gray-700 mb-4">過去のご利用履歴</h2>
        <div className="bg-white rounded-xl border border-dashed border-gray-300 p-10 text-center text-gray-400">
          <p className="text-sm">過去の利用履歴はありません</p>
        </div>
      </section>

      <div className="mt-8 text-center">
        <Link href="/" className="inline-flex items-center gap-2 text-[#1A5276] font-medium hover:underline">
          <ArrowRight size={16} /> 商品を探す
        </Link>
      </div>
    </div>
  )
}
