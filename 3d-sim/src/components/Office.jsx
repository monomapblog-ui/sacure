import { useRef } from 'react'
import { useFrame } from '@react-three/fiber'
import * as THREE from 'three'

// カラフルな床タイル
function FloorTiles() {
  const colors = ['#FFE0E0', '#E0F0FF', '#E0FFE8', '#FFF8E0', '#F0E0FF', '#FFE8F0']
  const tiles = []
  for (let x = -5; x <= 4; x++) {
    for (let z = -5; z <= 4; z++) {
      const ci = ((x + 5) * 10 + (z + 5)) % colors.length
      tiles.push(
        <mesh key={`${x}-${z}`} position={[x + 0.5, -0.01, z + 0.5]} receiveShadow>
          <boxGeometry args={[0.98, 0.02, 0.98]} />
          <meshStandardMaterial color={colors[ci]} roughness={0.4} />
        </mesh>
      )
    }
  }
  return <>{tiles}</>
}

// デスク
function Desk({ position, color = '#8B6914', large = false }) {
  const w = large ? 1.8 : 1.2
  const d = large ? 1.0 : 0.7
  return (
    <group position={position}>
      {/* 天板 */}
      <mesh position={[0, 0.38, 0]} castShadow receiveShadow>
        <boxGeometry args={[w, 0.07, d]} />
        <meshStandardMaterial color={large ? '#FFD700' : color} roughness={0.4} metalness={large ? 0.3 : 0} />
      </mesh>
      {/* 脚4本 */}
      {[[-w/2+0.08, -d/2+0.08], [w/2-0.08, -d/2+0.08], [-w/2+0.08, d/2-0.08], [w/2-0.08, d/2-0.08]].map(([lx, lz], i) => (
        <mesh key={i} position={[lx, 0.18, lz]} castShadow>
          <cylinderGeometry args={[0.04, 0.04, 0.36, 8]} />
          <meshStandardMaterial color="#555" />
        </mesh>
      ))}
      {/* PC モニター */}
      <mesh position={[0, 0.65, -d/2+0.1]} castShadow>
        <boxGeometry args={[0.5, 0.32, 0.03]} />
        <meshStandardMaterial color="#1a1a2e" />
      </mesh>
      <mesh position={[0, 0.52, -d/2+0.12]}>
        <boxGeometry args={[0.08, 0.08, 0.05]} />
        <meshStandardMaterial color="#333" />
      </mesh>
      {/* 画面の光 */}
      <mesh position={[0, 0.65, -d/2+0.08]}>
        <boxGeometry args={[0.44, 0.26, 0.01]} />
        <meshStandardMaterial color="#4FC3F7" emissive="#4FC3F7" emissiveIntensity={0.5} />
      </mesh>
    </group>
  )
}

// 観覧車（背景装飾）
function FerrisWheel({ position }) {
  const ref = useRef()
  useFrame((_, delta) => { if (ref.current) ref.current.rotation.z += delta * 0.3 })
  const gondolas = Array.from({ length: 8 }, (_, i) => i)
  return (
    <group position={position}>
      {/* 外輪 */}
      <mesh ref={ref}>
        <torusGeometry args={[1.4, 0.06, 8, 32]} />
        <meshStandardMaterial color="#FF6B6B" metalness={0.5} roughness={0.3} />
        {gondolas.map(i => {
          const angle = (i / 8) * Math.PI * 2
          return (
            <mesh key={i} position={[Math.cos(angle) * 1.4, Math.sin(angle) * 1.4, 0]}>
              <boxGeometry args={[0.22, 0.18, 0.12]} />
              <meshStandardMaterial color={['#FFD700','#0099D4','#27AE60','#E67E22','#9B59B6','#E74C3C','#1ABC9C','#F39C12'][i]} />
            </mesh>
          )
        })}
      </mesh>
      {/* 内輪スポーク */}
      {gondolas.map(i => {
        const angle = (i / 8) * Math.PI * 2
        return (
          <mesh key={i} rotation={[0, 0, angle]}>
            <boxGeometry args={[2.8, 0.03, 0.03]} />
            <meshStandardMaterial color="#ccc" />
          </mesh>
        )
      })}
      {/* 支柱 */}
      <mesh position={[-1, -1.8, 0]}>
        <cylinderGeometry args={[0.06, 0.08, 1.6, 8]} />
        <meshStandardMaterial color="#888" />
      </mesh>
      <mesh position={[1, -1.8, 0]}>
        <cylinderGeometry args={[0.06, 0.08, 1.6, 8]} />
        <meshStandardMaterial color="#888" />
      </mesh>
    </group>
  )
}

// カラフル壁
function Walls() {
  const wallData = [
    { pos: [0, 1.5, -5.5], rot: [0,0,0], w: 11, color: '#FFE4E1' },
    { pos: [0, 1.5, 5.5],  rot: [0, Math.PI, 0], w: 11, color: '#E1F5FE' },
    { pos: [-5.5, 1.5, 0], rot: [0, Math.PI/2, 0], w: 11, color: '#E8F5E9' },
    { pos: [5.5, 1.5, 0],  rot: [0, -Math.PI/2, 0], w: 11, color: '#FFF9C4' },
  ]
  return <>
    {wallData.map((w, i) => (
      <mesh key={i} position={w.pos} rotation={w.rot} receiveShadow>
        <boxGeometry args={[w.w, 3, 0.15]} />
        <meshStandardMaterial color={w.color} roughness={0.8} />
      </mesh>
    ))}
  </>
}

// 会社看板
function CompanySign() {
  return (
    <group position={[0, 2.8, -5.3]}>
      <mesh>
        <boxGeometry args={[4, 0.8, 0.1]} />
        <meshStandardMaterial color="#003D5C" />
      </mesh>
      <mesh position={[0, 0, 0.06]}>
        <boxGeometry args={[3.8, 0.6, 0.02]} />
        <meshStandardMaterial color="#FFD700" emissive="#FFD700" emissiveIntensity={0.3} />
      </mesh>
    </group>
  )
}

// 星・旗バナー
function Banner() {
  const colors = ['#FF6B6B', '#FFD700', '#0099D4', '#27AE60', '#E67E22', '#9B59B6']
  return (
    <group>
      {colors.map((c, i) => (
        <group key={i} position={[-5 + i * 2, 2.8, -5.2]}>
          <mesh>
            <sphereGeometry args={[0.1, 8, 8]} />
            <meshStandardMaterial color={c} emissive={c} emissiveIntensity={0.4} />
          </mesh>
        </group>
      ))}
      {/* ガーランドロープ */}
      <mesh position={[0, 2.8, -5.2]} rotation={[0, 0, 0]}>
        <cylinderGeometry args={[0.02, 0.02, 10, 4]} />
        <meshStandardMaterial color="#FFD700" />
      </mesh>
    </group>
  )
}

// 植物
function Plant({ position }) {
  return (
    <group position={position}>
      <mesh position={[0, 0.2, 0]}>
        <cylinderGeometry args={[0.15, 0.12, 0.4, 8]} />
        <meshStandardMaterial color="#8B4513" />
      </mesh>
      <mesh position={[0, 0.55, 0]}>
        <sphereGeometry args={[0.3, 8, 8]} />
        <meshStandardMaterial color="#27AE60" roughness={0.9} />
      </mesh>
      <mesh position={[-0.15, 0.65, 0.1]}>
        <sphereGeometry args={[0.18, 8, 8]} />
        <meshStandardMaterial color="#2ECC71" roughness={0.9} />
      </mesh>
    </group>
  )
}

// 会議テーブル（中央）
function MeetingTable() {
  return (
    <group position={[0, 0, 0.5]}>
      <mesh position={[0, 0.35, 0]} receiveShadow>
        <cylinderGeometry args={[1.0, 1.0, 0.06, 16]} />
        <meshStandardMaterial color="#F0E68C" roughness={0.3} />
      </mesh>
      <mesh position={[0, 0.18, 0]}>
        <cylinderGeometry args={[0.06, 0.06, 0.36, 8]} />
        <meshStandardMaterial color="#888" />
      </mesh>
    </group>
  )
}

export default function Office() {
  return (
    <group>
      {/* 床 */}
      <FloorTiles />
      <mesh position={[0, -0.02, 0]} receiveShadow>
        <boxGeometry args={[11, 0.04, 11]} />
        <meshStandardMaterial color="#ddd" />
      </mesh>

      {/* 壁 */}
      <Walls />
      <CompanySign />
      <Banner />

      {/* 各デスク */}
      <Desk position={[0, 0, 3]} large={true} />
      <Desk position={[-3.5, 0, -1]} color="#A0522D" />
      <Desk position={[3.5, 0, -1]} color="#6B8E23" />
      <Desk position={[0, 0, -3.5]} color="#8B4513" />

      {/* 中央会議テーブル */}
      <MeetingTable />

      {/* 植物 */}
      <Plant position={[-4.8, 0, -4.8]} />
      <Plant position={[4.8, 0, -4.8]} />
      <Plant position={[-4.8, 0, 4.5]} />
      <Plant position={[4.8, 0, 4.5]} />

      {/* 観覧車（コーナー装飾） */}
      <FerrisWheel position={[-4.5, 2.2, -4.5]} />

      {/* 照明（天井） */}
      {[[-3, 3, -3], [3, 3, -3], [0, 3, 1], [-3, 3, 3], [3, 3, 3]].map(([x, y, z], i) => (
        <group key={i} position={[x, y, z]}>
          <mesh>
            <sphereGeometry args={[0.15, 8, 8]} />
            <meshStandardMaterial color="#FFF9C4" emissive="#FFF176" emissiveIntensity={1} />
          </mesh>
          <pointLight intensity={0.6} distance={4} color="#FFF9C4" />
        </group>
      ))}
    </group>
  )
}
