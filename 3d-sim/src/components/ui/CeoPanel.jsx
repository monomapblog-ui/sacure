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
        <div style={{ color: '#567899', fontSize: 9 }}>{(file.size / 1024).toFixed(0)} KB</div>
      </div>
      <button onClick={onRemove} style={{ background: 'none', border: 'none', color: '#567899', cursor: 'pointer', fontSize: 14, padding: 0 }}>✕</button>
    </div>
  )
}

function CopyButton({ text }) {
  const [copied, setCopied] = useState(false)
  const handleCopy = () => {
    navigator.clipboard.writeText(text).then(() => {
      setCopied(true)
      setTimeout(() => setCopied(false), 1500)
    })
  }
  return (
    <button onClick={handleCopy} style={{
      background: copied ? 'rgba(39,174,96,0.2)' : 'rgba(255,255,255,0.08)',
      border: `1px solid ${copied ? 'rgba(39,174,96,0.4)' : 'rgba(255,255,255,0.15)'}`,
      color: copied ? '#27AE60' : '#8BCFE8',
      borderRadius: 4, padding: '2px 8px', fontSize: 9,
      cursor: 'pointer', fontFamily: 'inherit', flexShrink: 0,
    }}>
      {copied ? '✓ コピー済み' : 'コピー'}
    </button>
  )
}

function LogEntry({ entry }) {
  const [expanded, setExpanded] = useState(true)
  const isMurapee = entry.from === 'murapee'
  const isSystem  = entry.from === 'system'
  const isLoading = entry.loading
  const empColor  = COLOR[entry.from]
  const isLong    = entry.text && entry.text.length > 200

  if (isSystem) return (
    <div style={{
      background: entry.error ? 'rgba(231,76,60,0.12)' : 'rgba(255,255,255,0.04)',
      border: `1px solid ${entry.error ? 'rgba(231,76,60,0.3)' : 'rgba(255,255,255,0.06)'}`,
      borderRadius: 8, padding: '6px 10px',
    }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 3 }}>
        <span style={{ color: entry.error ? '#E74C3C' : '#27AE60', fontSize: 9 }}>
          {entry.error ? '⚠️ エラー' : '✅ システム'}
        </span>
        <span style={{ color: '#567899', fontSize: 9 }}>{entry.time}</span>
      </div>
      <div style={{ color: '#8BCFE8', fontSize: 11, lineHeight: 1.5 }}>{entry.text}</div>
    </div>
  )

  if (isMurapee) return (
    <div style={{
      background: 'rgba(0,61,92,0.6)',
      border: '1px solid rgba(255,215,0,0.25)',
      borderRadius: 8, padding: '6px 10px',
    }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 3 }}>
        <span style={{ color: '#FFD700', fontSize: 9, fontWeight: 'bold' }}>
          👑 むらぴー → {LABEL[entry.target] || '全員'}
        </span>
        <span style={{ color: '#567899', fontSize: 9 }}>{entry.time}</span>
      </div>
      <div style={{ color: '#fff', fontSize: 11, lineHeight: 1.5 }}>{entry.text}</div>
      {entry.file && (
        <div style={{ display: 'flex', alignItems: 'center', gap: 5, marginTop: 5, background: 'rgba(255,215,0,0.08)', borderRadius: 5, padding: '3px 7px' }}>
          <span style={{ fontSize: 12 }}>{entry.file.type?.startsWith('image/') ? '🖼️' : '📎'}</span>
          <span style={{ color: '#FFD700', fontSize: 9 }}>{entry.file.name}</span>
        </div>
      )}
    </div>
  )

  // 社員の返答（成果物）
  return (
    <div style={{
      background: isLoading ? 'rgba(255,255,255,0.03)' : 'rgba(255,255,255,0.06)',
      border: `1px solid ${empColor ? empColor + '44' : 'rgba(255,255,255,0.08)'}`,
      borderLeft: `3px solid ${empColor || '#567899'}`,
      borderRadius: '0 8px 8px 0', padding: '6px 10px',
      opacity: isLoading ? 0.7 : 1,
    }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 5 }}>
        <span style={{ color: empColor || '#8BCFE8', fontSize: 9, fontWeight: 'bold' }}>
          {LABEL[entry.from] || entry.name || entry.from}
        </span>
        <div style={{ display: 'flex', gap: 4, alignItems: 'center' }}>
          <span style={{ color: '#567899', fontSize: 9 }}>{entry.time}</span>
          {!isLoading && isLong && (
            <button onClick={() => setExpanded(v => !v)} style={{
              background: 'none', border: 'none', color: '#567899',
              cursor: 'pointer', fontSize: 9, padding: '0 2px',
            }}>
              {expanded ? '▲ 閉じる' : '▼ 開く'}
            </button>
          )}
          {!isLoading && <CopyButton text={entry.text} />}
        </div>
      </div>

      {isLoading ? (
        <div style={{ color: '#567899', fontSize: 11, display: 'flex', alignItems: 'center', gap: 6 }}>
          <span style={{ animation: 'pulse 1s infinite' }}>💭</span> 作成中...
        </div>
      ) : (
        <div style={{
          color: '#E8F4FC', fontSize: 11, lineHeight: 1.7,
          whiteSpace: 'pre-wrap', wordBreak: 'break-word',
          maxHeight: expanded ? 'none' : '60px',
          overflow: 'hidden',
          fontFamily: 'monospace',
        }}>
          {entry.text}
        </div>
      )}
    </div>
  )
}

export default function CeoPanel() {
  const { log, statuses, sendInstruction, resetHistory } = useSim()
  const [text, setText] = useState('')
  const [target, setTarget] = useState('all')
  const [attachedFile, setAttachedFile] = useState(null)
  const [sending, setSending] = useState(false)
  const logRef = useRef(null)
  const fileInputRef = useRef(null)

  useEffect(() => {
    if (logRef.current) logRef.current.scrollTop = 0
  }, [log])

  const handleSend = async () => {
    if ((!text.trim() && !attachedFile) || sending) return
    const msg = text.trim() || `📎 ${attachedFile.name}`
    setSending(true)
    try {
      await sendInstruction(msg, target, attachedFile ? { name: attachedFile.name, type: attachedFile.type } : null)
    } finally {
      setSending(false)
      setText('')
      setAttachedFile(null)
    }
  }

  const handleKey = (e) => {
    if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); handleSend() }
  }

  const handleFileChange = (e) => {
    const f = e.target.files[0]
    if (f) setAttachedFile(f)
    e.target.value = ''
  }

  const canSend = (text.trim() || attachedFile) && !sending

  return (
    <div style={{
      position: 'absolute', top: 0, right: 0,
      width: 280, height: '100vh',
      background: 'rgba(8, 16, 32, 0.95)',
      borderLeft: '2px solid #FFD700',
      display: 'flex', flexDirection: 'column',
      fontFamily: '"Hiragino Kaku Gothic ProN", "Noto Sans JP", sans-serif',
      backdropFilter: 'blur(12px)',
      zIndex: 10,
    }}>
      <style>{`
        @keyframes pulse { 0%,100%{opacity:1} 50%{opacity:.4} }
        @keyframes spin { from{transform:rotate(0deg)} to{transform:rotate(360deg)} }
      `}</style>

      {/* ヘッダー */}
      <div style={{
        background: 'linear-gradient(135deg, #003D5C, #005F8A)',
        borderBottom: '2px solid #FFD700',
        padding: '10px 12px',
        flexShrink: 0,
      }}>
        <div style={{ color: '#FFD700', fontWeight: 'bold', fontSize: 14, display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
          <span>👑 むらぴーの指示室</span>
          <button
            onClick={resetHistory}
            title="会話履歴をリセット"
            style={{ background: 'none', border: '1px solid rgba(255,255,255,.2)', color: 'rgba(255,255,255,.5)', borderRadius: 4, padding: '2px 7px', fontSize: 9, cursor: 'pointer', fontFamily: 'inherit' }}
          >
            リセット
          </button>
        </div>
        <div style={{ color: '#B0DFF2', fontSize: 9, marginTop: 2 }}>指示すると各AIが実際に返答します</div>
      </div>

      {/* 社員ステータス */}
      <div style={{ padding: '8px 10px', borderBottom: '1px solid rgba(255,255,255,0.1)', flexShrink: 0 }}>
        <div style={{ color: '#8BCFE8', fontSize: 9, fontWeight: 'bold', letterSpacing: '0.08em', marginBottom: 6 }}>■ ステータス</div>
        {EMPLOYEES.map(emp => {
          const s = statuses[emp.id]
          const isBusy = s.status !== 'working'
          return (
            <div key={emp.id} style={{
              display: 'flex', alignItems: 'center', gap: 7,
              padding: '4px 7px', borderRadius: 5, marginBottom: 3,
              background: isBusy ? 'rgba(255,215,0,0.08)' : 'rgba(255,255,255,0.03)',
              border: `1px solid ${isBusy ? 'rgba(255,215,0,0.3)' : 'rgba(255,255,255,0.06)'}`,
            }}>
              <div style={{
                width: 7, height: 7, borderRadius: '50%', background: emp.bodyColor, flexShrink: 0,
                boxShadow: isBusy ? `0 0 6px ${emp.bodyColor}` : 'none',
                animation: isBusy ? 'pulse 1s infinite' : 'none',
              }} />
              <div style={{ flex: 1 }}>
                <div style={{ color: '#fff', fontSize: 10, fontWeight: 'bold' }}>{emp.name}</div>
                <div style={{ color: isBusy ? '#FFD700' : '#567899', fontSize: 8, marginTop: 1 }}>
                  {isBusy ? '💭 むらぴーに返答中...' : `⚡ ${s.task}`}
                </div>
              </div>
            </div>
          )
        })}
      </div>

      {/* ログ */}
      <div ref={logRef} style={{
        flex: 1, overflowY: 'auto', padding: '8px 10px',
        display: 'flex', flexDirection: 'column', gap: 5,
      }}>
        <div style={{ color: '#8BCFE8', fontSize: 9, fontWeight: 'bold', letterSpacing: '0.08em', marginBottom: 2 }}>■ 会話ログ</div>
        {log.map(entry => <LogEntry key={entry.id} entry={entry} />)}
      </div>

      {/* 宛先選択 */}
      <div style={{ padding: '6px 10px', borderTop: '1px solid rgba(255,255,255,0.1)', flexShrink: 0 }}>
        <div style={{ color: '#8BCFE8', fontSize: 9, marginBottom: 5 }}>指示先</div>
        <div style={{ display: 'flex', gap: 4, flexWrap: 'wrap' }}>
          {['all', 'lp', 'sales', 'biz'].map(t => (
            <button key={t} onClick={() => setTarget(t)} style={{
              background: target === t ? COLOR[t] : 'rgba(255,255,255,0.07)',
              color: target === t ? '#000' : '#aaa',
              border: `1px solid ${target === t ? COLOR[t] : 'rgba(255,255,255,0.12)'}`,
              borderRadius: 10, padding: '2px 9px', fontSize: 9, fontWeight: 'bold',
              cursor: 'pointer', fontFamily: 'inherit',
            }}>
              {LABEL[t]}
            </button>
          ))}
        </div>
      </div>

      {/* 入力エリア */}
      <div style={{ padding: '6px 10px 12px', flexShrink: 0, borderTop: '1px solid rgba(255,255,255,0.1)' }}>
        {attachedFile && <FilePreview file={attachedFile} onRemove={() => setAttachedFile(null)} />}
        <div style={{ position: 'relative' }}>
          <textarea
            value={text}
            onChange={e => setText(e.target.value)}
            onKeyDown={handleKey}
            placeholder={sending ? '返答待ち中...' : '指示を入力... (Enter で送信)'}
            disabled={sending}
            rows={3}
            style={{
              width: '100%', resize: 'none',
              background: sending ? 'rgba(255,255,255,0.04)' : 'rgba(255,255,255,0.07)',
              border: `1px solid ${sending ? 'rgba(255,255,255,0.1)' : 'rgba(255,215,0,0.3)'}`,
              borderRadius: 7, padding: '7px 34px 7px 9px',
              color: '#fff', fontSize: 11,
              fontFamily: 'inherit', outline: 'none',
              boxSizing: 'border-box',
            }}
          />
          <button
            onClick={() => fileInputRef.current?.click()}
            disabled={sending}
            title="ファイルを添付"
            style={{
              position: 'absolute', right: 7, bottom: 9,
              background: 'none', border: 'none',
              color: attachedFile ? '#FFD700' : '#567899',
              cursor: sending ? 'default' : 'pointer', fontSize: 15, padding: 0,
            }}
          >📎</button>
          <input ref={fileInputRef} type="file"
            accept="image/*,.pdf,.doc,.docx,.xls,.xlsx,.txt"
            onChange={handleFileChange} style={{ display: 'none' }} />
        </div>
        <button onClick={handleSend} disabled={!canSend} style={{
          width: '100%', marginTop: 5,
          background: canSend ? 'linear-gradient(135deg, #003D5C, #0099D4)' : 'rgba(255,255,255,0.06)',
          color: canSend ? '#FFD700' : '#444',
          border: `2px solid ${canSend ? '#FFD700' : 'rgba(255,255,255,0.08)'}`,
          borderRadius: 7, padding: '7px', fontSize: 12,
          fontWeight: 'bold', fontFamily: 'inherit',
          cursor: canSend ? 'pointer' : 'default',
          transition: 'all 0.15s',
          display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 6,
        }}>
          {sending
            ? <><span style={{ animation: 'spin 1s linear infinite', display: 'inline-block' }}>⟳</span> 返答待ち中...</>
            : '👑 指示を送る'
          }
        </button>
      </div>
    </div>
  )
}
