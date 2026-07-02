import { NextResponse } from 'next/server'
import { z } from 'zod'

const registerSchema = z.object({
  name: z.string().min(1, 'お名前は必須です'),
  email: z.string().email('有効なメールアドレスを入力してください'),
  password: z.string().min(8, 'パスワードは8文字以上必要です'),
  phone: z.string().optional(),
  userType: z.enum(['INDIVIDUAL', 'CORPORATE']),
  companyName: z.string().optional(),
  department: z.string().optional(),
  corporateNo: z.string().optional(),
})

export async function POST(request: Request) {
  try {
    const body = await request.json()
    const data = registerSchema.parse(body)

    if (data.userType === 'CORPORATE' && !data.companyName) {
      return NextResponse.json(
        { error: '法人の場合は会社名が必須です' },
        { status: 400 }
      )
    }

    // TODO: DB保存 + bcryptjs でパスワードハッシュ化
    // const hash = await bcrypt.hash(data.password, 12)
    // const user = await prisma.user.create({ data: { ...data, passwordHash: hash } })

    return NextResponse.json(
      { message: '登録完了（モック）', userId: 'mock-user-id' },
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
