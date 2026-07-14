import { useRef, useState } from 'react'
import { useFrame } from '@react-three/fiber'
import { Html } from '@react-three/drei'
import * as THREE from 'three'

const BOSS_POS = new THREE.Vector3(0, 0, 3)

export function Boss({ data }) {
  const groupRef = useRef()
  const [task, setTask] = useState(data.tasks[0])
  let t = useRef(0)

  useFrame((state, delta) => {
    if (!groupRef.current) return
    t.current += delta
    const idx = Math.floor(t.current / 6) % data.tasks.length
    setTask(data.tasks[idx])
    // ゆっくり左右に揺れる
    groupRef.current.rotation.y = Math.sin(state.clock.elapsedTime * 0.5) * 0.3
    groupRef.current.position.y = Math.sin(state.clock.elapsedTime * 1.2) * 0.04
  })

  return (
    <group ref={groupRef} position={[data.desk[0], 0, data.desk[2]]}>
      {/* 体 */}
      <mesh position={[0, 0.65, 0]} castShadow>
        <capsuleGeometry args={[0.25, 0.55, 4, 8]} />
        <meshStandardMaterial color={data.bodyColor} roughness={0.6} />
      </mesh>
      {/* ネクタイ */}
      <mesh position={[0, 0.65, 0.24]} castShadow>
        <boxGeometry args={[0.1, 0.35, 0.02]} />
        <meshStandardMaterial color="#FFD700" />
      </mesh>
      {/* 頭 */}
      <mesh position={[0, 1.28, 0]} castShadow>
        <sphereGeometry args={[0.23, 16, 16]} />
        <meshStandardMaterial color="#FDBCB4" roughness={0.5} />
      </mesh>
      {/* 目（左右） */}
      <mesh position={[0.08, 1.31, 0.2]}>
        <sphereGeometry args={[0.035, 8, 8]} />
        <meshStandardMaterial color="#1a1a1a" />
      </mesh>
      <mesh position={[-0.08, 1.31, 0.2]}>
        <sphereGeometry args={[0.035, 8, 8]} />
        <meshStandardMaterial color="#1a1a1a" />
      </mesh>
      {/* 笑顔 */}
      <mesh position={[0, 1.22, 0.22]} rotation={[0.2, 0, 0]}>
        <torusGeometry args={[0.06, 0.012, 4, 8, Math.PI]} />
        <meshStandardMaterial color="#c0392b" />
      </mesh>
      {/* 王冠 */}
      <mesh position={[0, 1.58, 0]}>
        <cylinderGeometry args={[0.15, 0.22, 0.18, 5]} />
        <meshStandardMaterial color="#FFD700" metalness={0.8} roughness={0.2} />
      </mesh>
      <mesh position={[0, 1.68, 0]}>
        <sphereGeometry args={[0.07, 8, 8]} />
        <meshStandardMaterial color="#E74C3C" emissive="#E74C3C" emissiveIntensity={0.3} />
      </mesh>
      {[-0.13, 0, 0.13].map((x, i) => (
        <mesh key={i} position={[x, 1.64, 0]}>
          <sphereGeometry args={[0.04, 8, 8]} />
          <meshStandardMaterial color="#FFD700" metalness={0.9} />
        </mesh>
      ))}
      {/* オーラ（光リング） */}
      <mesh position={[0, 0, 0]} rotation={[Math.PI / 2, 0, 0]}>
        <torusGeometry args={[0.5, 0.04, 8, 32]} />
        <meshStandardMaterial color="#FFD700" emissive="#FFD700" emissiveIntensity={0.5} transparent opacity={0.4} />
      </mesh>
      {/* ラベル */}
      <Html position={[0, 2.1, 0]} center distanceFactor={10} zIndexRange={[100, 0]}>
        <div style={{
          background: 'linear-gradient(135deg, #003D5C, #0099D4)',
          color: '#fff',
          padding: '6px 14px',
          borderRadius: '20px',
          fontSize: '15px',
          fontWeight: 'bold',
          whiteSpace: 'nowrap',
          border: '2px solid #FFD700',
          boxShadow: '0 0 12px rgba(255,215,0,0.5)',
          fontFamily: 'sans-serif',
          textAlign: 'center',
        }}>
          👑 {data.name}
          <span style={{ display: 'block', fontSize: '10px', color: '#FFD700', fontWeight: 'normal' }}>
            {data.role}
          </span>
        </div>
        <div style={{
          background: 'rgba(255,215,0,0.15)',
          color: '#FFD700',
          padding: '3px 10px',
          borderRadius: '6px',
          fontSize: '11px',
          marginTop: '4px',
          whiteSpace: 'nowrap',
          fontFamily: 'sans-serif',
          textAlign: 'center',
          border: '1px solid rgba(255,215,0,0.3)',
        }}>
          {task}
        </div>
      </Html>
    </group>
  )
}

export function Employee({ data }) {
  const groupRef = useRef()
  const posRef = useRef(new THREE.Vector3(data.desk[0], 0, data.desk[2]))
  const targetRef = useRef(new THREE.Vector3(data.desk[0], 0, data.desk[2]))
  const stateRef = useRef('WORKING')
  const timerRef = useRef(data.reportInterval * (0.5 + Math.random()))
  const taskIdxRef = useRef(0)
  const [uiState, setUiState] = useState({ task: data.tasks[0], talking: false, walking: false })

  useFrame((state, delta) => {
    if (!groupRef.current) return

    timerRef.current -= delta

    // 状態遷移
    if (timerRef.current <= 0) {
      if (stateRef.current === 'WORKING') {
        stateRef.current = 'WALKING_TO_BOSS'
        const offset = new THREE.Vector3((Math.random() - 0.5) * 0.6, 0, 1.3)
        targetRef.current.copy(BOSS_POS).add(offset)
        setUiState(prev => ({ ...prev, walking: true, talking: false }))
      } else if (stateRef.current === 'TALKING') {
        stateRef.current = 'WALKING_BACK'
        targetRef.current.set(data.desk[0], 0, data.desk[2])
        taskIdxRef.current = (taskIdxRef.current + 1) % data.tasks.length
        timerRef.current = data.reportInterval + Math.random() * 5
        setUiState({ task: data.tasks[taskIdxRef.current], talking: false, walking: true })
      }
    }

    // 移動
    const dist = posRef.current.distanceTo(targetRef.current)
    if (dist > 0.08) {
      posRef.current.lerp(targetRef.current, Math.min(delta * 2.8, 1))
      const dir = new THREE.Vector3().subVectors(targetRef.current, posRef.current)
      if (dir.length() > 0.05) {
        groupRef.current.rotation.y = Math.atan2(dir.x, dir.z)
      }
      // 歩行ボブ
      groupRef.current.position.y = Math.abs(Math.sin(state.clock.elapsedTime * 8)) * 0.06
    } else {
      groupRef.current.position.y = 0
      if (stateRef.current === 'WALKING_TO_BOSS') {
        stateRef.current = 'TALKING'
        timerRef.current = 3 + Math.random() * 2
        setUiState(prev => ({ ...prev, talking: true, walking: false }))
        // むらぴー方向に向く
        const dir = new THREE.Vector3().subVectors(BOSS_POS, posRef.current)
        groupRef.current.rotation.y = Math.atan2(dir.x, dir.z)
      } else if (stateRef.current === 'WALKING_BACK') {
        stateRef.current = 'WORKING'
        setUiState(prev => ({ ...prev, walking: false }))
      }
    }

    // 作業中アニメ（ゆっくり首振り）
    if (stateRef.current === 'WORKING') {
      groupRef.current.rotation.y = Math.sin(state.clock.elapsedTime * 0.8 + data.desk[0]) * 0.15
    }

    groupRef.current.position.x = posRef.current.x
    groupRef.current.position.z = posRef.current.z
  })

  return (
    <group ref={groupRef} position={[data.desk[0], 0, data.desk[2]]}>
      {/* 体 */}
      <mesh position={[0, 0.62, 0]} castShadow>
        <capsuleGeometry args={[0.2, 0.45, 4, 8]} />
        <meshStandardMaterial color={data.bodyColor} roughness={0.6} />
      </mesh>
      {/* 頭 */}
      <mesh position={[0, 1.18, 0]} castShadow>
        <sphereGeometry args={[0.2, 16, 16]} />
        <meshStandardMaterial color="#FDBCB4" roughness={0.5} />
      </mesh>
      {/* 目 */}
      <mesh position={[0.07, 1.21, 0.17]}>
        <sphereGeometry args={[0.03, 8, 8]} />
        <meshStandardMaterial color="#1a1a1a" />
      </mesh>
      <mesh position={[-0.07, 1.21, 0.17]}>
        <sphereGeometry args={[0.03, 8, 8]} />
        <meshStandardMaterial color="#1a1a1a" />
      </mesh>
      {/* ラベル */}
      <Html position={[0, 1.85, 0]} center distanceFactor={10} zIndexRange={[100, 0]}>
        <div style={{
          background: data.bodyColor,
          color: '#fff',
          padding: '4px 12px',
          borderRadius: '16px',
          fontSize: '13px',
          fontWeight: 'bold',
          whiteSpace: 'nowrap',
          fontFamily: 'sans-serif',
          textAlign: 'center',
          boxShadow: '0 2px 8px rgba(0,0,0,0.3)',
        }}>
          {data.name}
          <span style={{ display: 'block', fontSize: '9px', color: data.accentColor, fontWeight: 'normal' }}>
            {data.role}
          </span>
        </div>
        {uiState.talking && (
          <div style={{
            background: '#fff',
            color: '#003D5C',
            padding: '4px 10px',
            borderRadius: '10px',
            fontSize: '11px',
            marginTop: '4px',
            border: '2px solid #FFD700',
            whiteSpace: 'nowrap',
            fontFamily: 'sans-serif',
            boxShadow: '0 2px 6px rgba(255,215,0,0.4)',
          }}>
            💬 指示受け中...
          </div>
        )}
        {!uiState.talking && (
          <div style={{
            background: 'rgba(0,0,0,0.6)',
            color: uiState.walking ? '#FFD700' : data.accentColor,
            padding: '2px 8px',
            borderRadius: '6px',
            fontSize: '10px',
            marginTop: '4px',
            whiteSpace: 'nowrap',
            fontFamily: 'sans-serif',
          }}>
            {uiState.walking ? '🚶 移動中...' : `⚡ ${uiState.task}`}
          </div>
        )}
      </Html>
    </group>
  )
}
