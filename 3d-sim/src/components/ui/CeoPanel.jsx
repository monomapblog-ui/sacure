import { useState, useRef, useEffect } from 'react'
import { useSim } from '../../context/SimContext'
import { EMPLOYEES } from '../../data'

const LABEL = { lp: 'LP制作君', sales: '営業企画君', biz: '経営企画君', all: '全員' }
const COLOR  = { lp: '#0099D4', sales: '#27AE60', biz: '#E67E22', all: '#FFD700' }

function FilePreview({ file, onRemove }) {
  const isImage = file.type.startsWith('image/')
  const [src, setSrc] = useState(null)

  useEffect(() => {
    if (isImage) {
      const reader = new FileReader()
      reader.onload = e => setSrc(e.target.result)
      reader.readAsDataURL(file)
    }
  }, [file, isImage])

  return (
    <div style={{
      display: 'flex', alignItems: 'center', gap: 6,
      background: 'rgba(255,215,0,0.1)',
      border: '1px solid rgba(255,215,0,0.3)',
      borderRadius: 6, padding: '4px 8px', marginBottom: 6,
    }}>
      {isImage && src
        ? <img src={src} alt="" style={{ width: 36, height: 36, objectFit: 'cover', borderRadius: 4 }} />
        : <span style={{ fontSize: 20 }}>📄</span>
      }
      <div style={{ flex: 1, overflow: 'hidden' }}>
        <div style={{ color: '#FFD700', fontSize: 10, fontWeight: 'bold', whiteSpace: 'nowrap', textOverflow: 'ellipsis', overflow: 'hidden' }}>
          {file.name}
        </div>
        <div style={{ color: '#567899', fontSize: 9 }}>
          {(file.size / 1024).toFixed(0)} KB
        </div>
      </div>
      <button onClick={onRemove} style={{
        background: 'none', border: 'none', color: '#567899',
        cursor: 'pointer', fontSize: 14, padding: 0, lineHeight: 1,
      }}>✕</button>
    </div>
  )
}

export default function CeoPanel() {
  const { log, statuses, sendInstruction } = useSim()
  const [text, setText] = useState('')
  const [target, setTarget] = useState('all')
  const [attachedFile, setAttachedFile] = useState(null)
  const logRef = useRef(null)
  const fileInputRef = useRef(null)

  useEffect(() => {
    if (logRef.current) logRef.current.scrollTop = 0
  }, [log])

  const handleSend = () => {
    if (!text.trim() && !attachedFile) return
    const msg = text.trim() || `📎 ${attachedFile.name}`
    sendInstruction(msg, target, attachedFile ? { name: attachedFile.name, type: attachedFile.type } : null)
    setText('')
    setAttachedFile(null)
  }

  const handleKey = (e) => {
    if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); handleSend() }
  }

  const handleFileChange = (e) => {
    const f = e.target.files[0]
    if (f) setAttachedFile(f)
    e.target.value = ''
  }

  const canSend = text.trim() || attachedFile

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
        <div style={{ color: '#B0DFF2', fontSize: 10, marginTop: 2 }}>社員に指示・資料を送れます</div>
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
            <div style={{ color: entry.from === 'murapee' ? '#fff' : '#8BCFE8', fontSize: 11, lineHeight: 1.5 }}>
              {entry.text}
            </div>
            {entry.file && (
              <div style={{
                display: 'flex', alignItems: 'center', gap: 5,
                marginTop: 5, background: 'rgba(255,215,0,0.08)',
                borderRadius: 5, padding: '3px 7px',
              }}>
                <span style={{ fontSize: 12 }}>{entry.file.type?.startsWith('image/') ? '🖼️' : '📎'}</span>
                <span style={{ color: '#FFD700', fontSize: 9 }}>{entry.file.name}</span>
              </div>
            )}
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
        {attachedFile && (
          <FilePreview file={attachedFile} onRemove={() => setAttachedFile(null)} />
        )}
        <div style={{ position: 'relative' }}>
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
              borderRadius: 8, padding: '8px 36px 8px 10px',
              color: '#fff', fontSize: 12,
              fontFamily: 'inherit', outline: 'none',
              boxSizing: 'border-box',
            }}
          />
          {/* ファイル添付ボタン */}
          <button
            onClick={() => fileInputRef.current?.click()}
            title="ファイルを添付"
            style={{
              position: 'absolute', right: 8, bottom: 10,
              background: 'none', border: 'none',
              color: attachedFile ? '#FFD700' : '#567899',
              cursor: 'pointer', fontSize: 16, padding: 0, lineHeight: 1,
            }}
          >
            📎
          </button>
          <input
            ref={fileInputRef}
            type="file"
            accept="image/*,.pdf,.doc,.docx,.xls,.xlsx,.ppt,.pptx,.txt"
            onChange={handleFileChange}
            style={{ display: 'none' }}
          />
        </div>
        <button onClick={handleSend} style={{
          width: '100%', marginTop: 6,
          background: canSend ? 'linear-gradient(135deg, #003D5C, #0099D4)' : 'rgba(255,255,255,0.08)',
          color: canSend ? '#FFD700' : '#555',
          border: `2px solid ${canSend ? '#FFD700' : 'rgba(255,255,255,0.1)'}`,
          borderRadius: 8, padding: '8px', fontSize: 13,
          fontWeight: 'bold', fontFamily: 'inherit',
          cursor: canSend ? 'pointer' : 'default',
          transition: 'all 0.2s',
        }}>
          👑 指示を送る
        </button>
      </div>
    </div>
  )
}
