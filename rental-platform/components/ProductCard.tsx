import Link from 'next/link'
import { Package } from 'lucide-react'
import { Product } from '@/types'
import { formatPrice, getStatusLabel } from '@/lib/utils'

interface Props {
  product: Product
}

const CATEGORY_EMOJI: Record<string, string> = {
  '冷蔵庫': '🧊',
  '洗濯機': '🫧',
  '電子レンジ': '📦',
  'その他家電': '⚡',
}

export default function ProductCard({ product }: Props) {
  const available = product.status === 'AVAILABLE'

  return (
    <div className={`bg-white rounded-xl shadow-md overflow-hidden border-2 transition-transform hover:-translate-y-1 hover:shadow-lg ${available ? 'border-transparent' : 'border-gray-200 opacity-75'}`}>
      {/* 画像エリア */}
      <div className="bg-gray-100 h-44 flex items-center justify-center relative">
        <span className="text-6xl">{CATEGORY_EMOJI[product.category] ?? '📦'}</span>
        {/* グレードバッジ */}
        <span className={`absolute top-3 left-3 text-xs font-bold px-2 py-1 rounded-full ${product.grade === 'A' ? 'bg-green-500 text-white' : 'bg-yellow-500 text-white'}`}>
          {product.grade}ランク
        </span>
        {/* ステータス */}
        {!available && (
          <div className="absolute inset-0 bg-black/40 flex items-center justify-center">
            <span className="bg-gray-800 text-white text-sm font-bold px-4 py-2 rounded-full">
              {getStatusLabel(product.status)}
            </span>
          </div>
        )}
      </div>

      <div className="p-4">
        <p className="text-xs text-gray-400 mb-1">{product.category} / {product.brand}</p>
        <h3 className="font-bold text-gray-800 text-base mb-2 leading-snug">{product.name}</h3>
        <p className="text-xs text-gray-500 mb-3 line-clamp-2">{product.description}</p>

        <div className="flex items-end justify-between">
          <div>
            <span className="text-2xl font-bold text-[#1A5276]">{formatPrice(product.monthlyPrice)}</span>
            <span className="text-xs text-gray-400 ml-1">/月</span>
          </div>
          {available ? (
            <Link href={`/products/${product.id}`}
              className="bg-[#1E8B4C] hover:bg-green-600 text-white text-sm font-bold px-4 py-2 rounded-lg transition-colors">
              詳細・申込
            </Link>
          ) : (
            <span className="text-xs text-gray-400 bg-gray-100 px-3 py-2 rounded-lg">
              {getStatusLabel(product.status)}
            </span>
          )}
        </div>
      </div>
    </div>
  )
}
