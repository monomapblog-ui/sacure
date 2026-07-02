'use client'
import { use } from 'react'
import { notFound } from 'next/navigation'
import Link from 'next/link'
import { MOCK_PRODUCTS } from '@/lib/mock-data'
import { formatPrice, formatDate, getGradeLabel } from '@/lib/utils'
import { ArrowLeft, CheckCircle, Clock, MapPin, Wrench } from 'lucide-react'

export default function ProductDetailPage({ params }: { params: Promise<{ id: string }> }) {
  const { id } = use(params)
  const product = MOCK_PRODUCTS.find(p => p.id === id)
  if (!product) notFound()

  const available = product.status === 'AVAILABLE'

  return (
    <div className="max-w-4xl mx-auto px-4 py-10">
      <Link href="/" className="flex items-center gap-2 text-gray-500 hover:text-[#1A5276] mb-6 text-sm">
        <ArrowLeft size={16} /> 商品一覧に戻る
      </Link>

      <div className="bg-white rounded-2xl shadow-md overflow-hidden">
        {/* 画像エリア */}
        <div className="bg-gray-100 h-64 flex items-center justify-center">
          <span className="text-8xl">
            {{'冷蔵庫':'🧊','洗濯機':'🫧','電子レンジ':'📦','その他家電':'⚡'}[product.category] ?? '📦'}
          </span>
        </div>

        <div className="p-8">
          <div className="flex flex-wrap items-start justify-between gap-4 mb-6">
            <div>
              <span className="text-xs text-gray-400">{product.category} / {product.brand}</span>
              <h1 className="text-2xl font-bold text-gray-800 mt-1">{product.name}</h1>
              <span className={`inline-block mt-2 text-xs font-bold px-3 py-1 rounded-full ${product.grade === 'A' ? 'bg-green-100 text-green-700' : 'bg-yellow-100 text-yellow-700'}`}>
                {getGradeLabel(product.grade)}
              </span>
            </div>
            <div className="text-right">
              <p className="text-3xl font-bold text-[#1A5276]">{formatPrice(product.monthlyPrice)}<span className="text-base font-normal text-gray-400">/月</span></p>
              <p className="text-xs text-gray-400 mt-1">最低契約3ヶ月 / 自動更新</p>
            </div>
          </div>

          <p className="text-gray-600 mb-8 leading-relaxed">{product.description}</p>

          {/* 循環トレーサビリティ */}
          <div className="bg-green-50 border border-green-200 rounded-xl p-5 mb-8">
            <h3 className="font-bold text-green-800 mb-4 flex items-center gap-2">
              ♻ この商品の来歴
            </h3>
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
              <div className="flex items-start gap-3">
                <MapPin size={18} className="text-green-600 mt-0.5 flex-shrink-0" />
                <div>
                  <p className="text-xs text-gray-500">回収場所</p>
                  <p className="text-sm font-medium text-gray-700">{product.sourceLocation ?? '東京都内'}</p>
                </div>
              </div>
              <div className="flex items-start gap-3">
                <Wrench size={18} className="text-green-600 mt-0.5 flex-shrink-0" />
                <div>
                  <p className="text-xs text-gray-500">整備完了日</p>
                  <p className="text-sm font-medium text-gray-700">{product.refurbishedAt ? formatDate(product.refurbishedAt) : '—'}</p>
                </div>
              </div>
              <div className="flex items-start gap-3">
                <CheckCircle size={18} className="text-green-600 mt-0.5 flex-shrink-0" />
                <div>
                  <p className="text-xs text-gray-500">品質検査</p>
                  <p className="text-sm font-medium text-gray-700">100点満点チェック済み</p>
                </div>
              </div>
            </div>
          </div>

          {/* 料金・条件 */}
          <div className="border border-gray-200 rounded-xl p-5 mb-8">
            <h3 className="font-bold text-gray-700 mb-4">料金・契約条件</h3>
            <div className="grid grid-cols-2 gap-4 text-sm">
              {[
                ['月額料金', formatPrice(product.monthlyPrice)],
                ['最低契約期間', '3ヶ月'],
                ['途中解約', '残存期間の50〜100%'],
                ['延滞料', '月額の1/30（日割り）'],
                ['配送・設置', '無料（東京23区・神奈川）'],
                ['返却送料', '無料（集荷対応）'],
              ].map(([k, v]) => (
                <div key={k} className="flex justify-between border-b border-gray-100 pb-2">
                  <span className="text-gray-500">{k}</span>
                  <span className="font-medium text-gray-800">{v}</span>
                </div>
              ))}
            </div>
          </div>

          {/* 申込ボタン */}
          {available ? (
            <Link href="/auth/register"
              className="block w-full bg-[#1E8B4C] hover:bg-green-600 text-white text-center font-bold text-lg py-4 rounded-xl transition-colors">
              この商品を申し込む
            </Link>
          ) : (
            <div className="w-full bg-gray-200 text-gray-500 text-center font-bold text-lg py-4 rounded-xl flex items-center justify-center gap-2">
              <Clock size={20} /> 現在貸出中です
            </div>
          )}
          <p className="text-center text-xs text-gray-400 mt-3">
            申込にはアカウント登録（無料）が必要です
          </p>
        </div>
      </div>
    </div>
  )
}
