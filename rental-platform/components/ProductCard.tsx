import Link from 'next/link'
import { Product } from '@/types'
import { getStatusLabel } from '@/lib/utils'

interface Props {
  product: Product
}

const CATEGORY_EMOJI: Record<string, string> = {
  '冷蔵庫': '🧊',
  '洗濯機': '🫧',
  '電子レンジ': '📦',
  '3点セット': '📦',
}

export default function ProductCard({ product }: Props) {
  const available = product.status === 'AVAILABLE'
  const usedPrice = product.monthlyPriceUsed ?? Math.floor(product.monthlyPrice * 0.66)

  return (
    <div className={`bg-white rounded-2xl shadow-md overflow-hidden border transition-transform hover:-translate-y-1 hover:shadow-xl ${available ? 'border-gray-100' : 'border-gray-200 opacity-70'}`}>
      {/* 画像エリア */}
      <div className="bg-gradient-to-br from-[#EBF6FC] to-[#D0EEF9] h-44 flex items-center justify-center relative">
        <span className="text-7xl">{CATEGORY_EMOJI[product.category] ?? '📦'}</span>
        {!available && (
          <div className="absolute inset-0 bg-black/40 flex items-center justify-center">
            <span className="bg-gray-800 text-white text-sm font-bold px-4 py-2 rounded-full">
              {getStatusLabel(product.status)}
            </span>
          </div>
        )}
      </div>

      <div className="p-5">
        <h3 className="font-bold text-gray-800 text-lg mb-1">{product.name}</h3>
        <p className="text-xs text-gray-500 mb-4 leading-relaxed">{product.description}</p>

        {/* 価格表示 */}
        <div className="bg-gray-50 rounded-xl p-3 mb-4 space-y-2">
          <div className="flex items-center justify-between">
            <span className="text-xs font-bold text-gray-500 bg-gray-200 px-2 py-0.5 rounded">新品</span>
            <div className="flex items-baseline gap-0.5">
              <span className="text-xl font-bold text-[#0099D4]">¥{product.monthlyPrice.toLocaleString()}</span>
              <span className="text-xs text-gray-400">/月</span>
            </div>
          </div>
          <div className="flex items-center justify-between">
            <span className="text-xs font-bold text-gray-500 bg-amber-100 text-amber-700 px-2 py-0.5 rounded">中古</span>
            <div className="flex items-baseline gap-0.5">
              <span className="text-xl font-bold text-amber-600">¥{usedPrice.toLocaleString()}</span>
              <span className="text-xs text-gray-400">/月</span>
            </div>
          </div>
        </div>

        {available ? (
          <Link href={`/products/${product.id}`}
            className="block w-full text-center bg-[#0099D4] hover:bg-[#007BAA] text-white text-sm font-bold py-2.5 rounded-xl transition-colors">
            詳細・申込 →
          </Link>
        ) : (
          <div className="text-center text-xs text-gray-400 bg-gray-100 py-2.5 rounded-xl">
            {getStatusLabel(product.status)}
          </div>
        )}
      </div>
    </div>
  )
}
