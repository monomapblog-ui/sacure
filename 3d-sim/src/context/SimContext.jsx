import { createContext, useContext, useState, useCallback, useRef } from 'react'

const SimContext = createContext(null)

const BACKEND = 'http://localhost:3002'

const INITIAL_STATUSES = {
  lp:    { task: 'LPをデザイン中', status: 'working', instruction: null },
  sales: { task: '外回り準備中',   status: 'working', instruction: null },
  biz:   { task: '経営戦略を立案中', status: 'working', instruction: null },
}

const EMP_COLORS = {
  lp:    '#0099D4',
  sales: '#27AE60',
  biz:   '#E67E22',
}

const EMP_NAMES = {
  lp:    'LP制作君',
  sales: '営業企画君',
  biz:   '経営企画君',
}

export function SimProvider({ children }) {
  const [log, setLog] = useState([
    { id: 0, from: 'system', text: 'AIシミュレーター起動しました。右パネルからむらぴーとして指示を送れます。', time: '00:00' },
  ])
  const [statuses, setStatuses] = useState(INITIAL_STATUSES)
  const [activeInstruction, setActiveInstruction] = useState(null)
  const idRef = useRef(1)

  const nextId = () => ++idRef.current

  const sendInstruction = useCallback(async (text, target, file = null) => {
    const now = new Date().toLocaleTimeString('ja-JP', { hour: '2-digit', minute: '2-digit' })
    const targets = target === 'all' ? ['lp', 'sales', 'biz'] : [target]
    const instructionId = nextId()

    // むらぴーの指示をログに追加
    setLog(prev => [{
      id: instructionId, from: 'murapee',
      text, time: now, target, file,
    }, ...prev])

    setActiveInstruction({ text, target })

    // 対象社員を召喚状態に
    setStatuses(prev => {
      const next = { ...prev }
      targets.forEach(t => {
        next[t] = { ...next[t], status: 'summoned', instruction: text }
      })
      return next
    })

    // 「考え中...」ローディングエントリを追加
    const loadingIds = {}
    targets.forEach(t => {
      const lid = nextId()
      loadingIds[t] = lid
      setLog(prev => [{
        id: lid, from: t,
        text: '考え中...', time: now, loading: true,
      }, ...prev])
    })

    try {
      const res = await fetch(`${BACKEND}/api/chat`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ instruction: text, target }),
      })

      if (!res.ok) throw new Error(`サーバーエラー: ${res.status}`)
      const data = await res.json()
      const doneTime = new Date().toLocaleTimeString('ja-JP', { hour: '2-digit', minute: '2-digit' })

      // ローディングを削除して返答を挿入
      setLog(prev => {
        const filtered = prev.filter(e => !Object.values(loadingIds).includes(e.id))
        const replies = data.responses.map(r => ({
          id: nextId(), from: r.target, name: r.name,
          text: r.response, time: doneTime,
        }))
        return [...replies, ...filtered]
      })

    } catch (err) {
      console.error(err)
      const errTime = new Date().toLocaleTimeString('ja-JP', { hour: '2-digit', minute: '2-digit' })
      setLog(prev => {
        const filtered = prev.filter(e => !Object.values(loadingIds).includes(e.id))
        return [{
          id: nextId(), from: 'system',
          text: `⚠️ エラー: ${err.message} — .envにAPIキーが設定されているか確認してください`,
          time: errTime, error: true,
        }, ...filtered]
      })
    } finally {
      // 社員を通常業務に戻す
      setStatuses(prev => {
        const next = { ...prev }
        targets.forEach(t => {
          next[t] = { ...next[t], status: 'working', instruction: null }
        })
        return next
      })
      setActiveInstruction(null)
    }
  }, [])

  const resetHistory = useCallback(async () => {
    try {
      await fetch(`${BACKEND}/api/reset`, { method: 'POST' })
      const now = new Date().toLocaleTimeString('ja-JP', { hour: '2-digit', minute: '2-digit' })
      setLog([{ id: nextId(), from: 'system', text: '会話履歴をリセットしました', time: now }])
    } catch {
      // backend not running
    }
  }, [])

  const updateStatus = useCallback((id, patch) => {
    setStatuses(prev => ({ ...prev, [id]: { ...prev[id], ...patch } }))
  }, [])

  return (
    <SimContext.Provider value={{
      log, statuses, activeInstruction,
      sendInstruction, updateStatus, resetHistory,
      EMP_COLORS, EMP_NAMES,
    }}>
      {children}
    </SimContext.Provider>
  )
}

export const useSim = () => useContext(SimContext)
