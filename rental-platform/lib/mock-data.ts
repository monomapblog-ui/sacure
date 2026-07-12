import { Product, Order } from '@/types'

export const MOCK_PRODUCTS: Product[] = [
  {
    id: 'set001',
    name: '家電3点セット',
    category: '3点セット',
    grade: 'A',
    status: 'AVAILABLE',
    monthlyPrice: 2980,
    monthlyPriceUsed: 1980,
    description: '冷蔵庫（100〜150L）＋ 洗濯機（4.5kg）＋ 電子レンジ。配送・設置・退去時回収まですべて込み。',
    createdAt: '2026-01-01',
    isSet: true,
    setItems: ['冷蔵庫（100〜150L）', '洗濯機（4.5kg）', '電子レンジ（700W）'],
  },
  {
    id: 'p001',
    name: '冷蔵庫',
    category: '冷蔵庫',
    grade: 'A',
    status: 'AVAILABLE',
    monthlyPrice: 2980,
    monthlyPriceUsed: 1980,
    description: '100〜150L・2ドア冷蔵庫。冷凍・冷蔵ともに清潔。動作確認・クリーニング済み。',
    createdAt: '2026-01-01',
  },
  {
    id: 'p002',
    name: '洗濯機',
    category: '洗濯機',
    grade: 'A',
    status: 'AVAILABLE',
    monthlyPrice: 2980,
    monthlyPriceUsed: 1980,
    description: '4.5kg全自動洗濯機。槽洗浄・パッキン清掃済み。静音設計で賃貸にも安心。',
    createdAt: '2026-01-01',
  },
  {
    id: 'p003',
    name: '電子レンジ',
    category: '電子レンジ',
    grade: 'A',
    status: 'AVAILABLE',
    monthlyPrice: 2980,
    monthlyPriceUsed: 1980,
    description: '700W単機能電子レンジ。フラットテーブルで洗いやすい。クリーニング・動作確認済み。',
    createdAt: '2026-01-01',
  },
]

export const MOCK_ORDERS: Order[] = [
  {
    id: 'ord001',
    userId: 'u001',
    productId: 'set001',
    status: 'ACTIVE',
    startDate: '2026-05-01',
    minEndDate: '2026-08-01',
    monthlyPrice: 2980,
    deliveryAddress: '東京都渋谷区渋谷3-6-7',
    createdAt: '2026-04-28',
    product: MOCK_PRODUCTS.find(p => p.id === 'set001'),
  },
]

export const CATEGORIES = ['すべて', '3点セット', '冷蔵庫', '洗濯機', '電子レンジ']
