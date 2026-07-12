'use client'
import Link from 'next/link'
import ProductCard from '@/components/ProductCard'
import { MOCK_PRODUCTS } from '@/lib/mock-data'
import { RefreshCw, Shield, Truck, CheckCircle } from 'lucide-react'

const SET = MOCK_PRODUCTS.find(p => p.isSet)!
const SINGLES = MOCK_PRODUCTS.filter(p => !p.isSet)

export default function HomePage() {
  return (
    <>
      {/* ヒーロー */}
      <section className="bg-[#003D5C] text-white py-16 px-4">
        <div className="max-w-4xl mx-auto text-center">
          <p className="text-[#FFD700] text-sm font-bold mb-3 tracking-widest">♻ CIRCULAR ECONOMY RENTAL</p>
          <h1 className="text-4xl md:text-5xl font-bold mb-4 leading-tight">
            使い終わった家電に、<br />もう一度命を。
          </h1>
          <p className="text-gray-300 text-lg mb-8">
            サキュレが回収・再生した家電を月額レンタルでお届け。<br />
            初期費用ゼロ・最短3ヶ月から・配送〜回収まですべて込み。
          </p>
          <div className="flex flex-wrap justify-center gap-3 text-sm">
            {[
              { icon: <RefreshCw size={15}/>, text: '回収→再生→レンタルの循環モデル' },
              { icon: <Shield size={15}/>, text: '新品・中古どちらも品質保証付き' },
              { icon: <Truck size={15}/>, text: '東京23区・神奈川県 配送無料' },
            ].map((item, i) => (
              <span key={i} className="flex items-center gap-2 bg-white/10 px-4 py-2 rounded-full">
                {item.icon}{item.text}
              </span>
            ))}
          </div>
        </div>
      </section>

      {/* 3点セット */}
      <section className="max-w-5xl mx-auto px-4 py-14">
        <div className="flex items-center gap-3 mb-6">
          <span className="w-1 h-7 bg-[#0099D4] rounded-full block"/>
          <h2 className="text-2xl font-bold text-[#003D5C]">留学生・新生活スタートパック</h2>
          <span className="bg-[#FFD700] text-[#003D5C] text-xs font-bold px-3 py-1 rounded-full">人気No.1</span>
        </div>

        <div className="bg-gradient-to-r from-[#003D5C] to-[#0099D4] rounded-2xl overflow-hidden shadow-xl">
          <div className="grid md:grid-cols-2 gap-0">
            {/* 左: 商品情報 */}
            <div className="p-8 text-white">
              <div className="text-sm font-bold text-[#FFD700] mb-3 tracking-widest">3点セット</div>
              <div className="flex gap-6 mb-6">
                {[
                  { emoji: '🧊', label: '冷蔵庫\n100〜150L' },
                  { emoji: '🫧', label: '洗濯機\n4.5kg' },
                  { emoji: '📦', label: '電子レンジ\n700W' },
                ].map((item, i) => (
                  <div key={i} className="text-center">
                    <div className="text-4xl mb-2">{item.emoji}</div>
                    <div className="text-xs text-blue-200 whitespace-pre-line leading-snug">{item.label}</div>
                  </div>
                ))}
              </div>
              <div className="space-y-2">
                {[
                  '配送・設置・退去時回収 すべて無料',
                  '新品・中古どちらも品質チェック済み',
                  '最低3ヶ月〜・途中解約相談可',
                  '英語・中国語サポート対応',
                ].map((text, i) => (
                  <div key={i} className="flex items-center gap-2 text-sm text-blue-100">
                    <CheckCircle size={14} className="text-[#FFD700] flex-shrink-0"/>
                    {text}
                  </div>
                ))}
              </div>
            </div>

            {/* 右: 価格 */}
            <div className="bg-white p-8 flex flex-col justify-center">
              <div className="text-xs font-bold text-gray-400 mb-4 tracking-widest">月額レンタル料金</div>

              <div className="space-y-4 mb-6">
                <div className="flex items-center justify-between bg-[#EBF6FC] rounded-xl px-5 py-4">
                  <div>
                    <span className="text-xs font-bold text-[#0099D4] block mb-0.5">新品</span>
                    <span className="text-xs text-gray-500">品質保証・動作確認済み</span>
                  </div>
                  <div className="text-right">
                    <span className="text-3xl font-bold text-[#003D5C]">¥2,980</span>
                    <span className="text-sm text-gray-400">/月</span>
                  </div>
                </div>

                <div className="flex items-center justify-between bg-amber-50 rounded-xl px-5 py-4">
                  <div>
                    <span className="text-xs font-bold text-amber-600 block mb-0.5">中古（リファービッシュ）</span>
                    <span className="text-xs text-gray-500">回収・再生品・品質チェック済み</span>
                  </div>
                  <div className="text-right">
                    <span className="text-3xl font-bold text-amber-600">¥1,980</span>
                    <span className="text-sm text-gray-400">/月</span>
                  </div>
                </div>
              </div>

              <Link href={`/products/${SET.id}`}
                className="block w-full text-center bg-[#0099D4] hover:bg-[#007BAA] text-white font-bold py-3.5 rounded-xl text-base transition-colors">
                3点セットを申し込む →
              </Link>
              <p className="text-center text-xs text-gray-400 mt-3">東京23区・神奈川県 配送・設置・回収無料</p>
            </div>
          </div>
        </div>
      </section>

      {/* 単体レンタル */}
      <section className="max-w-5xl mx-auto px-4 pb-14">
        <div className="flex items-center gap-3 mb-6">
          <span className="w-1 h-7 bg-[#0099D4] rounded-full block"/>
          <h2 className="text-2xl font-bold text-[#003D5C]">単体レンタル</h2>
          <span className="text-sm text-gray-400">必要な家電だけ選べます</span>
        </div>

        <div className="grid grid-cols-1 sm:grid-cols-3 gap-6">
          {SINGLES.map(product => (
            <ProductCard key={product.id} product={product} />
          ))}
        </div>

        <div className="mt-6 bg-[#F8FCFF] border border-[#D0EEF9] rounded-xl p-4 text-sm text-[#003D5C] text-center">
          ※ 単体でも3点セット同様、配送・設置・退去回収は無料です。複数台ご希望の場合はお気軽にご相談ください。
        </div>
      </section>

      {/* 特長 */}
      <section className="bg-white py-16 px-4 border-t">
        <div className="max-w-5xl mx-auto">
          <h2 className="text-2xl font-bold text-center text-[#003D5C] mb-10">サキュレレンタルが選ばれる理由</h2>
          <div className="grid grid-cols-1 md:grid-cols-3 gap-8">
            {[
              { icon: '♻️', title: '循環型で環境に優しい', desc: '廃棄されるはずだった家電を回収・再生してお届け。1台のご利用が廃棄物削減に直結します。' },
              { icon: '💰', title: '新品¥2,980・中古¥1,980/月', desc: '初期費用ゼロ。配送・設置・保証・回収まですべて月額料金に含まれています。' },
              { icon: '🛡️', title: '品質保証で安心', desc: '100点満点の品質チェックをパスした商品のみ。新品・中古どちらも動作保証付き。' },
            ].map((item, i) => (
              <div key={i} className="text-center p-6 rounded-2xl border border-gray-100 hover:shadow-md transition-shadow">
                <div className="text-4xl mb-4">{item.icon}</div>
                <h3 className="font-bold text-[#003D5C] mb-2">{item.title}</h3>
                <p className="text-gray-600 text-sm leading-relaxed">{item.desc}</p>
              </div>
            ))}
          </div>
        </div>
      </section>
    </>
  )
}
