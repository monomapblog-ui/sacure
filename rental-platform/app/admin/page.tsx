import Link from 'next/link'
import { MOCK_PRODUCTS, MOCK_ORDERS } from '@/lib/mock-data'
import { formatPrice } from '@/lib/utils'

export default function AdminDashboard() {
  const available = MOCK_PRODUCTS.filter(p => p.status === 'AVAILABLE').length
  const rented = MOCK_PRODUCTS.filter(p => p.status === 'RENTED').length
  const maintenance = MOCK_PRODUCTS.filter(p => p.status === 'MAINTENANCE').length
  const monthlyRevenue = MOCK_ORDERS.filter(o => o.status === 'ACTIVE').reduce((s, o) => s + o.monthlyPrice, 0)

  const stats = [
    { label: '在庫中', value: available, color: 'bg-green-500', unit: '台' },
    { label: '貸出中', value: rented, color: 'bg-blue-500', unit: '台' },
    { label: '整備中', value: maintenance, color: 'bg-yellow-500', unit: '台' },
    { label: '月次売上（予測）', value: formatPrice(monthlyRevenue), color: 'bg-[#1A5276]', unit: '' },
  ]

  return (
    <div className="max-w-6xl mx-auto px-4 py-10">
      <div className="flex items-center justify-between mb-8">
        <div>
          <h1 className="text-2xl font-bold text-[#1A5276]">管理ダッシュボード</h1>
          <p className="text-gray-500 text-sm mt-1">スーパー管理者ビュー</p>
        </div>
        <div className="flex gap-3">
          <Link href="/admin/products" className="bg-[#1A5276] hover:bg-blue-900 text-white text-sm font-bold px-4 py-2 rounded-lg transition-colors">
            商品管理
          </Link>
          <Link href="/admin/orders" className="bg-[#1E8B4C] hover:bg-green-600 text-white text-sm font-bold px-4 py-2 rounded-lg transition-colors">
            注文管理
          </Link>
        </div>
      </div>

      {/* KPI */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-10">
        {stats.map(s => (
          <div key={s.label} className="bg-white rounded-xl shadow-sm p-5 border border-gray-100">
            <div className={`w-3 h-3 rounded-full ${s.color} mb-3`} />
            <p className="text-xs text-gray-400">{s.label}</p>
            <p className="text-2xl font-bold text-gray-800 mt-1">{s.value}{s.unit}</p>
          </div>
        ))}
      </div>

      {/* 最近の在庫 */}
      <section className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
        <div className="px-6 py-4 border-b border-gray-100 flex items-center justify-between">
          <h2 className="font-bold text-gray-700">在庫一覧（直近6件）</h2>
          <Link href="/admin/products" className="text-xs text-[#1A5276] hover:underline">すべて見る →</Link>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead className="bg-gray-50">
              <tr>
                {['商品名', 'カテゴリ', 'ランク', '月額', 'ステータス', '回収場所'].map(h => (
                  <th key={h} className="px-4 py-3 text-left text-xs font-medium text-gray-500">{h}</th>
                ))}
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {MOCK_PRODUCTS.map(p => (
                <tr key={p.id} className="hover:bg-gray-50">
                  <td className="px-4 py-3 font-medium text-gray-800">{p.name}</td>
                  <td className="px-4 py-3 text-gray-500">{p.category}</td>
                  <td className="px-4 py-3">
                    <span className={`text-xs font-bold px-2 py-0.5 rounded-full ${p.grade === 'A' ? 'bg-green-100 text-green-700' : 'bg-yellow-100 text-yellow-700'}`}>
                      {p.grade}
                    </span>
                  </td>
                  <td className="px-4 py-3 text-gray-700">{formatPrice(p.monthlyPrice)}</td>
                  <td className="px-4 py-3">
                    <span className={`text-xs font-bold px-2 py-0.5 rounded-full ${
                      p.status === 'AVAILABLE' ? 'bg-green-100 text-green-700' :
                      p.status === 'RENTED' ? 'bg-blue-100 text-blue-700' :
                      'bg-yellow-100 text-yellow-700'
                    }`}>
                      {{AVAILABLE:'在庫中', RENTED:'貸出中', MAINTENANCE:'整備中', RETIRED:'廃棄'}[p.status]}
                    </span>
                  </td>
                  <td className="px-4 py-3 text-gray-500 text-xs">{p.sourceLocation ?? '—'}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>
    </div>
  )
}
