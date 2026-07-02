'use client'
import { useState } from 'react'
import ProductCard from '@/components/ProductCard'
import { MOCK_PRODUCTS, CATEGORIES } from '@/lib/mock-data'
import { RefreshCw, Shield, Truck } from 'lucide-react'

export default function HomePage() {
  const [category, setCategory] = useState('すべて')
  const [gradeFilter, setGradeFilter] = useState('すべて')

  const filtered = MOCK_PRODUCTS.filter(p => {
    if (category !== 'すべて' && p.category !== category) return false
    if (gradeFilter === 'A' && p.grade !== 'A') return false
    if (gradeFilter === 'B' && p.grade !== 'B') return false
    return true
  })

  return (
    <>
      {/* ヒーローセクション */}
      <section className="bg-[#1A5276] text-white py-16 px-4">
        <div className="max-w-4xl mx-auto text-center">
          <p className="text-green-300 text-sm font-bold mb-3 tracking-widest">♻ CIRCULAR ECONOMY RENTAL</p>
          <h1 className="text-4xl md:text-5xl font-bold mb-4 leading-tight">
            使い終わった家電に、<br />もう一度命を。
          </h1>
          <p className="text-gray-300 text-lg mb-8">
            サキュレが回収・再生した家電を月額レンタルでお届け。<br />
            初期費用ゼロ・最短3ヶ月から・法人・個人どちらも対応。
          </p>
          <div className="flex flex-wrap justify-center gap-4 text-sm">
            {[
              { icon: <RefreshCw size={16}/>, text: '回収→再生→レンタルの循環モデル' },
              { icon: <Shield size={16}/>, text: '品質保証付き（Aランク・Bランク）' },
              { icon: <Truck size={16}/>, text: '東京23区・神奈川県対応' },
            ].map((item, i) => (
              <span key={i} className="flex items-center gap-2 bg-white/10 px-4 py-2 rounded-full">
                {item.icon}{item.text}
              </span>
            ))}
          </div>
        </div>
      </section>

      {/* 商品一覧 */}
      <section className="max-w-6xl mx-auto px-4 py-10">
        <div className="flex flex-wrap gap-3 mb-8">
          <div className="flex flex-wrap gap-2">
            {CATEGORIES.map(cat => (
              <button key={cat}
                onClick={() => setCategory(cat)}
                className={`px-4 py-2 rounded-full text-sm font-medium transition-colors ${category === cat ? 'bg-[#1A5276] text-white' : 'bg-white text-gray-600 border hover:border-[#1A5276]'}`}>
                {cat}
              </button>
            ))}
          </div>
          <div className="flex gap-2 ml-auto">
            {['すべて', 'A', 'B'].map(g => (
              <button key={g}
                onClick={() => setGradeFilter(g)}
                className={`px-4 py-2 rounded-full text-sm font-medium transition-colors ${gradeFilter === g ? 'bg-green-600 text-white' : 'bg-white text-gray-600 border hover:border-green-600'}`}>
                {g === 'すべて' ? '全ランク' : `${g}ランク`}
              </button>
            ))}
          </div>
        </div>

        <p className="text-gray-500 text-sm mb-6">{filtered.length}件の商品</p>

        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-6">
          {filtered.map(product => (
            <ProductCard key={product.id} product={product} />
          ))}
        </div>
      </section>

      {/* 特長セクション */}
      <section className="bg-white py-16 px-4 mt-8">
        <div className="max-w-5xl mx-auto">
          <h2 className="text-2xl font-bold text-center text-[#1A5276] mb-10">サキュレレンタルが選ばれる理由</h2>
          <div className="grid grid-cols-1 md:grid-cols-3 gap-8">
            {[
              { icon: '♻️', title: '循環型で環境に優しい', desc: '廃棄されるはずだった家電を再生。1台レンタルするごとに、廃棄物削減に貢献できます。' },
              { icon: '💰', title: '新品の10〜15%の月額', desc: 'リファービッシュ品だから実現できる低価格。初期費用ゼロで始められます。' },
              { icon: '🔧', title: '品質保証で安心', desc: '100点満点の品質チェックをパスした商品のみ。Aランク・Bランクで状態を明示。' },
            ].map((item, i) => (
              <div key={i} className="text-center p-6 rounded-xl border border-gray-100 hover:shadow-md transition-shadow">
                <div className="text-4xl mb-4">{item.icon}</div>
                <h3 className="font-bold text-[#1A5276] mb-2">{item.title}</h3>
                <p className="text-gray-600 text-sm leading-relaxed">{item.desc}</p>
              </div>
            ))}
          </div>
        </div>
      </section>
    </>
  )
}
