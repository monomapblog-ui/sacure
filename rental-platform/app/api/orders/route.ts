import { NextResponse } from 'next/server'
import { MOCK_ORDERS } from '@/lib/mock-data'
import { z } from 'zod'

export async function GET(request: Request) {
  const { searchParams } = new URL(request.url)
  const status = searchParams.get('status')

  let orders = [...MOCK_ORDERS]
  if (status) {
    orders = orders.filter(o => o.status === status)
  }

  return NextResponse.json({ orders, total: orders.length })
}

const createOrderSchema = z.object({
  productId: z.string().min(1),
  userId: z.string().min(1),
  deliveryAddress: z.string().min(1, '配送先住所は必須です'),
})

export async function POST(request: Request) {
  try {
    const body = await request.json()
    const data = createOrderSchema.parse(body)

    // TODO: Stripe サブスクリプション作成 + DB保存
    // const subscription = await stripe.subscriptions.create(...)
    // const order = await prisma.order.create({ data: { ...data, stripeSubscriptionId: subscription.id } })

    const startDate = new Date().toISOString()
    const minEndDate = new Date(Date.now() + 90 * 24 * 60 * 60 * 1000).toISOString()

    return NextResponse.json(
      {
        message: '申込を受け付けました（モック）',
        order: {
          id: 'mock-order-' + Date.now(),
          productId: data.productId,
          userId: data.userId,
          status: 'ACTIVE',
          startDate,
          minEndDate,
        },
      },
      { status: 201 }
    )
  } catch (error) {
    if (error instanceof z.ZodError) {
      return NextResponse.json(
        { error: error.issues[0]?.message ?? 'バリデーションエラー' },
        { status: 400 }
      )
    }
    return NextResponse.json({ error: 'サーバーエラーが発生しました' }, { status: 500 })
  }
}
