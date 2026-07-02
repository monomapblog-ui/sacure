import { NextResponse } from 'next/server'
import { MOCK_PRODUCTS } from '@/lib/mock-data'

export async function GET(request: Request) {
  const { searchParams } = new URL(request.url)
  const category = searchParams.get('category')
  const grade = searchParams.get('grade')
  const status = searchParams.get('status')

  let products = [...MOCK_PRODUCTS]

  if (category && category !== 'すべて') {
    products = products.filter(p => p.category === category)
  }
  if (grade && grade !== 'すべて') {
    products = products.filter(p => p.grade === grade)
  }
  if (status) {
    products = products.filter(p => p.status === status)
  }

  return NextResponse.json({ products, total: products.length })
}
