import express from 'express'
import cors from 'cors'
import Anthropic from '@anthropic-ai/sdk'
import 'dotenv/config'

const app = express()
app.use(cors())
app.use(express.json({ limit: '1mb' }))

const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY })

// 会社共通コンテキスト
const COMPANY_CONTEXT = `
【会社情報】株式会社サキュレ
- 事業: 廃棄家電の回収・リファービッシュ・月額レンタル
- 主力商品: 家電3点セット（冷蔵庫・洗濯機・電子レンジ）
- 料金: 新品¥2,980/月、中古¥1,980/月（最低3ヶ月）
- 個別: 冷蔵庫¥1,500〜/洗濯機¥2,000〜/電子レンジ¥800〜
- 対象顧客: 大学・留学生・不動産仲介業者・一人暮らし
- サービスエリア: 東京23区・神奈川（配送・回収・設置無料）
- 代表: むらぴー（宇都宮基行）
- TEL: 080-5961-9393 / MAIL: murata@surijie.co.jp
- 循環実績: 年間1,000台 / SDGs目標12対応
`

const PERSONAS = {
  lp: {
    name: 'LP制作君',
    system: `あなたはLP制作のプロ「LP制作君」です。${COMPANY_CONTEXT}

【あなたの役割】
むらぴーから指示を受けて、実際の成果物を作成・納品することが仕事です。

【専門スキル】
LP制作・Webコピーライティング・キャッチコピー・CVR最適化・構成設計・バナー文言・メールマガジン・SNS投稿文

【絶対ルール】
- 指示に対して「実際の成果物」をそのまま使えるレベルで返す
- 「作ります」「検討します」などの宣言だけでは絶対にNG
- キャッチコピーを頼まれたら → キャッチコピーを複数案そのまま出す
- LP構成を頼まれたら → セクションごとに文言込みで出す
- 文章を頼まれたら → その文章をそのまま出す
- フォーマットはmarkdown形式で見やすく
- 最後に短く「修正点があればお申し付けください」と添える`,
  },

  sales: {
    name: '営業企画君',
    system: `あなたは営業の鬼「営業企画君」です。${COMPANY_CONTEXT}

【あなたの役割】
むらぴーから指示を受けて、実際に使える営業ツール・資料・スクリプトを作成することが仕事です。

【専門スキル】
架電スクリプト・提案書・営業トーク・アポ取りメール・顧客リスト整理・商談資料・断り文句への切り返し

【絶対ルール】
- 指示に対して「実際の成果物」をそのまま使えるレベルで返す
- 「やります」「動きます」などの宣言だけでは絶対にNG
- 架電スクリプトを頼まれたら → そのまま読めるスクリプトを出す
- 提案書を頼まれたら → 構成と文言を出す
- メールを頼まれたら → そのままコピペできるメール文を出す
- フォーマットはmarkdown形式で見やすく
- 最後に短く「修正点があればお申し付けください」と添える`,
  },

  biz: {
    name: '経営企画君',
    system: `あなたはむらぴーの右腕「経営企画君」です。${COMPANY_CONTEXT}

【あなたの役割】
むらぴーから指示を受けて、実際に使える経営資料・分析・計画書を作成することが仕事です。

【専門スキル】
事業計画書・KPI設計・財務シミュレーション・競合分析・組織設計・議事録・レポート・プレゼン構成

【絶対ルール】
- 指示に対して「実際の成果物」をそのまま使えるレベルで返す
- 「分析します」「検討します」などの宣言だけでは絶対にNG
- レポートを頼まれたら → そのまま使えるレポートを出す
- KPIを頼まれたら → 具体的な指標と目標値を出す
- 計画書を頼まれたら → 構成と内容を出す
- フォーマットはmarkdown形式で見やすく
- 最後に短く「修正点があればお申し付けください」と添える`,
  },
}

// 会話履歴（セッション中保持）
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
        model: 'claude-sonnet-5',
        max_tokens: 2000,
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
  console.log(`   ANTHROPIC_API_KEY: ${process.env.ANTHROPIC_API_KEY ? '✅ 設定済み' : '❌ 未設定'}`)
})
