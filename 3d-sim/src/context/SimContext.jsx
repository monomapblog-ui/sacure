import { createContext, useContext, useState, useCallback } from 'react'

const SimContext = createContext(null)

const INITIAL_STATUSES = {
  lp:    { task: 'LPデザイン中', status: 'working', instruction: null },
  sales: { task: '提案書作成中', status: 'working', instruction: null },
  biz:   { task: '事業計画作成中', status: 'working', instruction: null },
}

export function SimProvider({ children }) {
  const [log, setLog] = useState([
    { id: 0, from: 'system', text: 'シミュレーター起動しました', time: '00:00' },
  ])
  const [statuses, setStatuses] = useState(INITIAL_STATUSES)
  const [activeInstruction, setActiveInstruction] = useState(null)

  const sendInstruction = useCallback((text, target) => {
    const now = new Date().toLocaleTimeString('ja-JP', { hour: '2-digit', minute: '2-digit' })
    const targets = target === 'all' ? ['lp', 'sales', 'biz'] : [target]
    const id = Date.now()

    setLog(prev => [{ id, from: 'murapee', text, time: now, target }, ...prev])
    setActiveInstruction({ text, target })

    // 社員を召喚状態に
    setStatuses(prev => {
      const next = { ...prev }
      targets.forEach(t => {
        next[t] = { ...next[t], status: 'summoned', instruction: text }
      })
      return next
    })

    // 一定時間後に完了
    setTimeout(() => {
      const doneTime = new Date().toLocaleTimeString('ja-JP', { hour: '2-digit', minute: '2-digit' })
      setLog(prev => [{ id: id + 1, from: 'system', text: `「${text.slice(0, 20)}...」対応完了`, time: doneTime }, ...prev])
      setStatuses(prev => {
        const next = { ...prev }
        targets.forEach(t => {
          next[t] = { ...next[t], status: 'working', instruction: null }
        })
        return next
      })
      setActiveInstruction(null)
    }, 10000)
  }, [])

  const updateStatus = useCallback((id, patch) => {
    setStatuses(prev => ({ ...prev, [id]: { ...prev[id], ...patch } }))
  }, [])

  return (
    <SimContext.Provider value={{ log, statuses, activeInstruction, sendInstruction, updateStatus }}>
      {children}
    </SimContext.Provider>
  )
}

export const useSim = () => useContext(SimContext)
