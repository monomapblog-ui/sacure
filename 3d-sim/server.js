import express from 'express'
import cors from 'cors'
import Anthropic from '@anthropic-ai/sdk'
import 'dotenv/config'

const app = express()
app.use(cors())
app.use(express.json({ limit: '1mb' }))

const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY })

const PERSONAS = {
  lp: {
    name: 'LP制作君',
    system: `あなたはLP制作のプロ「LP制作君」です。株式会社サキュレ（廃棄家電のリユース・レンタルサービス）に勤めています。
代表「むらぴー」から指示を受けて仕事をする優秀な部下です。

【専門スキル】LP制作・Webデザイン・コピーライティング・CVR最適化・バナー制作・UI/UX

【性格】クリエイティブで完璧主義。デザインとコピーへのこだわりが人一倍強い。仕事を振られると目を輝かせる。

【返答ルール】
- むらぴーの指示に対して、具体的にどう取り掛かるか・何を提案するかを答える
- 200文字以内で簡潔に
- 語尾は「〜します！」「〜ですね！」など元気よく
- 絵文字を2〜3個使う
- 返答の最後に「着手します！」など行動宣言を入れる`,
  },
  sales: {
    name: '営業企画君',
    system: `あなたは営業の鬼「営業企画君」です。株式会社サキュレ（廃棄家電のリユース・レンタルサービス）に勤めています。
代表「むらぴー」から指示を受けて仕事をする最強の営業マンです。

【専門スキル】新規開拓・提案書作成・商談クロージング・顧客管理・営業戦略・架電リスト作成

【性格】積極的でポジティブ。数字への執念が異常に強い。絶対に諦めない。むらぴーへの忠誠心が高い。

【返答ルール】
- むらぴーの指示に対して、具体的なアクションプランや数値目標を答える
- 200文字以内で簡潔に
- 語尾は「〜します！」「〜やります！」など強気に
- 絵文字を2〜3個使う（📊💪🔥など）
- 必ず「動きます！」「やります！」など力強い行動宣言で締める`,
  },
  biz: {
    name: '経営企画君',
    system: `あなたはむらぴーの右腕「経営企画君」です。株式会社サキュレ（廃棄家電のリユース・レンタルサービス）に勤めています。
代表「むらぴー」から指示を受けて経営全般を支える参謀役です。

【専門スキル】事業計画策定・KPI分析・財務計画・経営戦略・組織マネジメント・M&A・資金調達

【性格】論理的で分析的。むらぴーの意図を正確に汲み取り、一歩先の提案ができる。冷静だが会社への情熱は誰より強い。

【返答ルール】
- むらぴーの指示に対して、戦略的観点・リスク・優先度を踏まえて答える
- 200文字以内で簡潔に
- 落ち着いたトーンで、でも熱意を感じさせる
- 絵文字を1〜2個使う（📈📋など）
- 「確認・分析・提言します」など参謀らしい締め方をする`,
  },
}

// 各社員の会話履歴
const histories = { lp: [], sales: [], biz: [] }

app.post('/api/chat', async (req, res) => {
  const { instruction, target } = req.body
  if (!instruction) return res.status(400).json({ error: '指示が空です' })

  const targets = target === 'all' ? ['lp', 'sales', 'biz'] : [target]

  try {
    const responses = await Promise.all(targets.map(async (t) => {
      const persona = PERSONAS[t]
      histories[t].push({ role: 'user', content: instruction })

      const msg = await client.messages.create({
        model: 'claude-haiku-4-5-20251001',
        max_tokens: 300,
        system: persona.system,
        messages: histories[t],
      })

      const reply = msg.content[0].text
      histories[t].push({ role: 'assistant', content: reply })

      return { target: t, name: persona.name, response: reply }
    }))

    res.json({ responses })
  } catch (err) {
    console.error('API Error:', err.message)
    res.status(500).json({ error: err.message })
  }
})

// 会話履歴リセット
app.post('/api/reset', (_req, res) => {
  histories.lp = []
  histories.sales = []
  histories.biz = []
  res.json({ ok: true })
})

app.get('/api/health', (_req, res) => res.json({ ok: true }))

const PORT = process.env.PORT || 3002
app.listen(PORT, () => {
  console.log(`\n🚀 サキュレAIバックエンド起動: http://localhost:${PORT}`)
  console.log(`   ANTHROPIC_API_KEY: ${process.env.ANTHROPIC_API_KEY ? '✅ 設定済み' : '❌ 未設定 (.envを確認)'}`)
})
