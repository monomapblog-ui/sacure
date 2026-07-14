import { Canvas } from '@react-three/fiber'
import { OrbitControls } from '@react-three/drei'
import { Suspense } from 'react'
import Office from './components/Office'
import { Boss, Employee } from './components/Employee'
import { BOSS, EMPLOYEES } from './data'
import { SimProvider } from './context/SimContext'
import CeoPanel from './components/ui/CeoPanel'

export default function App() {
  return (
    <SimProvider>
    <div style={{ width: '100vw', height: '100vh', background: '#1a0a2e', position: 'relative' }}>
      <Canvas shadows camera={{ position: [0, 14, 8], fov: 45 }} gl={{ antialias: true }}>
        <Suspense fallback={null}>
          <ambientLight intensity={0.5} />
          <directionalLight
            position={[8, 12, 6]} intensity={1.2} castShadow
            shadow-mapSize={[2048, 2048]}
            shadow-camera-left={-10} shadow-camera-right={10}
            shadow-camera-top={10} shadow-camera-bottom={-10}
          />
          <pointLight position={[0, 6, 3]} intensity={0.6} color="#FFD700" />
          <pointLight position={[-4, 4, -4]} intensity={0.4} color="#0099D4" />
          <pointLight position={[4, 4, -4]} intensity={0.4} color="#27AE60" />

          <Office />
          <Boss data={BOSS} />
          {EMPLOYEES.map(emp => <Employee key={emp.id} data={emp} />)}

          <OrbitControls
            target={[0, 0, 0]}
            minPolarAngle={Math.PI / 5}
            maxPolarAngle={Math.PI / 2.8}
            minDistance={8} maxDistance={20}
            enablePan={false}
          />
        </Suspense>
      </Canvas>

      {/* UIオーバーレイ */}
      <div style={{ position: 'absolute', top: 20, left: 20, fontFamily: '"Hiragino Kaku Gothic ProN", sans-serif' }}>
        <div style={{
          background: 'rgba(0,61,92,0.88)', border: '2px solid #FFD700',
          borderRadius: '12px', padding: '12px 20px', backdropFilter: 'blur(8px)',
        }}>
          <div style={{ color: '#FFD700', fontWeight: 'bold', fontSize: 18 }}>株式会社サキュレ</div>
          <div style={{ color: '#B0DFF2', fontSize: 12, marginTop: 2 }}>3D オフィスシミュレーター</div>
          <div style={{ marginTop: 10, display: 'flex', flexDirection: 'column', gap: 5 }}>
            {[
              { color: '#FFD700', name: 'むらぴー', role: '代表' },
              { color: '#0099D4', name: 'LP制作君', role: 'LP制作' },
              { color: '#27AE60', name: '営業企画君', role: '営業企画' },
              { color: '#E67E22', name: '経営企画君', role: '経営企画' },
            ].map(m => (
              <div key={m.name} style={{ display: 'flex', alignItems: 'center', gap: 7 }}>
                <div style={{ width: 10, height: 10, borderRadius: '50%', background: m.color, flexShrink: 0 }} />
                <span style={{ color: '#fff', fontSize: 12, fontWeight: 'bold' }}>{m.name}</span>
                <span style={{ color: '#8BCFE8', fontSize: 10 }}>{m.role}</span>
              </div>
            ))}
          </div>
        </div>
      </div>

      <div style={{ position: 'absolute', bottom: 16, left: 16, color: 'rgba(255,255,255,0.35)', fontSize: 11, fontFamily: 'sans-serif' }}>
        ドラッグで回転 / スクロールでズーム
      </div>

      <CeoPanel />
    </div>
    </SimProvider>
  )
}
