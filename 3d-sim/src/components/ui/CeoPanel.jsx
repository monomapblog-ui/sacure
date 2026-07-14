import { useState, useRef, useEffect } from 'react'
import { useSim } from '../../context/SimContext'
import { EMPLOYEES } from '../../data'

const LABEL = { lp: 'LP制作君', sales: '営業企画君', biz: '経営企画君', all: '全員' }
const COLOR  = { lp: '#0099D4', sales: '#27AE60', biz: '#E67E22', all: '#FFD700' }

export default function CeoPanel() {
  const { log, statuses, activeInstruction, sendInstruction } = useSim()
  const [text, setText] = useState('')
  const [target, setTarget] = useState('all')
  const logRef = useRef(null)

  useEffect(() => {
    if (logRef.current) logRef.current.scrollTop = 0
  }, [log])

  const handleSend = () => {
    if (!text.trim()) return
    sendInstruction(text.trim(), target)
    setText('')
  }

  const handleKey = (e) => {
    if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); handleSend() }
  }

  return (
    <div style={{
      position: 'absolute', top: 0, right: 0,
      width: 270, height: '100vh',
      background: 'rgba(8, 16, 32, 0.93)',
      borderLeft: '2px solid #FFD700',
      display: 'flex', flexDirection: 'column',
      fontFamily: '"Hiragino Kaku Gothic ProN", "Noto Sans JP", sans-serif',
      backdropFilter: 'blur(12px)',
      zIndex: 10,
    }}>

      {/* ヘッダー */}
      <div style={{
        background: 'linear-gradient(135deg, #003D5C, #005F8A)',
        borderBottom: '2px solid #FFD700',
        padding: '12px 14px',
        flexShrink: 0,
      }}>
        <div style={{ color: '#FFD700', fontWeight: 'bold', fontSize: 15, display: 'flex', alignItems: 'center', gap: 6 }}>
          👑 むらぴーの指示室
        </div>
        <div style={{ color: '#B0DFF2', fontSize: 10, marginTop: 2 }}>社員に指示を送れます</div>
      </div>

      {/* 社員ステータス */}
      <div style={{ padding: '10px 12px', borderBottom: '1px solid rgba(255,255,255,0.1)', flexShrink: 0 }}>
        <div style={{ color: '#8BCFE8', fontSize: 9, fontWeight: 'bold', letterSpacing: '0.08em', marginBottom: 8 }}>
          ■ 現在のステータス
        </div>
        {EMPLOYEES.map(emp => {
          const s = statuses[emp.id]
          const isBusy = s.status !== 'working'
          return (
            <div key={emp.id} style={{
              display: 'flex', alignItems: 'center', gap: 8,
              padding: '5px 8px', borderRadius: 6, marginBottom: 4,
              background: isBusy ? 'rgba(255,215,0,0.08)' : 'rgba(255,255,255,0.04)',
              border: `1px solid ${isBusy ? 'rgba(255,215,0,0.3)' : 'rgba(255,255,255,0.08)'}`,
            }}>
              <div style={{ width: 8, height: 8, borderRadius: '50%', background: emp.bodyColor, flexShrink: 0,
                boxShadow: isBusy ? `0 0 6px ${emp.bodyColor}` : 'none' }} />
              <div style={{ flex: 1 }}>
                <div style={{ color: '#fff', fontSize: 11, fontWeight: 'bold' }}>{emp.name}</div>
                <div style={{ color: isBusy ? '#FFD700' : '#567899', fontSize: 9, marginTop: 1 }}>
                  {isBusy ? '💬 むらぴーと対話中' : `⚡ ${s.task}`}
                </div>
              </div>
            </div>
          )
        })}
      </div>

      {/* ログ */}
      <div ref={logRef} style={{
        flex: 1, overflowY: 'auto', padding: '10px 12px',
        display: 'flex', flexDirection: 'column', gap: 6,
      }}>
        <div style={{ color: '#8BCFE8', fontSize: 9, fontWeight: 'bold', letterSpacing: '0.08em', marginBottom: 4 }}>
          ■ 指示ログ
        </div>
        {log.map(entry => (
          <div key={entry.id} style={{
            background: entry.from === 'murapee' ? 'rgba(0,61,92,0.6)' : 'rgba(255,255,255,0.04)',
            border: `1px solid ${entry.from === 'murapee' ? 'rgba(255,215,0,0.25)' : 'rgba(255,255,255,0.06)'}`,
            borderRadius: 8, padding: '6px 10px',
          }}>
            {entry.from === 'murapee' && (
              <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 3 }}>
                <span style={{ color: '#FFD700', fontSize: 9, fontWeight: 'bold' }}>
                  👑 むらぴー → {LABEL[entry.target] || '全員'}
                </span>
                <span style={{ color: '#567899', fontSize: 9 }}>{entry.time}</span>
              </div>
            )}
            {entry.from === 'system' && (
              <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 3 }}>
                <span style={{ color: '#27AE60', fontSize: 9 }}>✅ システム</span>
                <span style={{ color: '#567899', fontSize: 9 }}>{entry.time}</span>
              </div>
            )}
            <div style={{
              color: entry.from === 'murapee' ? '#fff' : '#8BCFE8',
              fontSize: 11, lineHeight: 1.5,
            }}>
              {entry.text}
            </div>
          </div>
        ))}
      </div>

      {/* 宛先選択 */}
      <div style={{ padding: '8px 12px', borderTop: '1px solid rgba(255,255,255,0.1)', flexShrink: 0 }}>
        <div style={{ color: '#8BCFE8', fontSize: 9, marginBottom: 6 }}>指示先</div>
        <div style={{ display: 'flex', gap: 4, flexWrap: 'wrap' }}>
          {['all', 'lp', 'sales', 'biz'].map(t => (
            <button key={t} onClick={() => setTarget(t)} style={{
              background: target === t ? COLOR[t] : 'rgba(255,255,255,0.08)',
              color: target === t ? '#000' : '#aaa',
              border: `1px solid ${target === t ? COLOR[t] : 'rgba(255,255,255,0.15)'}`,
              borderRadius: 12, padding: '3px 10px', fontSize: 10, fontWeight: 'bold',
              cursor: 'pointer', fontFamily: 'inherit',
            }}>
              {LABEL[t]}
            </button>
          ))}
        </div>
      </div>

      {/* 入力エリア */}
      <div style={{ padding: '8px 12px 14px', flexShrink: 0, borderTop: '1px solid rgba(255,255,255,0.1)' }}>
        <textarea
          value={text}
          onChange={e => setText(e.target.value)}
          onKeyDown={handleKey}
          placeholder="指示を入力... (Enter で送信)"
          rows={3}
          style={{
            width: '100%', resize: 'none',
            background: 'rgba(255,255,255,0.07)',
            border: '1px solid rgba(255,215,0,0.3)',
            borderRadius: 8, padding: '8px 10px',
            color: '#fff', fontSize: 12,
            fontFamily: 'inherit', outline: 'none',
            boxSizing: 'border-box',
          }}
        />
        <button onClick={handleSend} style={{
          width: '100%', marginTop: 6,
          background: text.trim() ? 'linear-gradient(135deg, #003D5C, #0099D4)' : 'rgba(255,255,255,0.08)',
          color: text.trim() ? '#FFD700' : '#555',
          border: `2px solid ${text.trim() ? '#FFD700' : 'rgba(255,255,255,0.1)'}`,
          borderRadius: 8, padding: '8px', fontSize: 13,
          fontWeight: 'bold', fontFamily: 'inherit',
          cursor: text.trim() ? 'pointer' : 'default',
          transition: 'all 0.2s',
        }}>
          👑 指示を送る
        </button>
      </div>
    </div>
  )
}
