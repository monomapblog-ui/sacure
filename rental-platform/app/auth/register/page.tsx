'use client'
import { useState } from 'react'
import Link from 'next/link'
import { UserType } from '@/types'

export default function RegisterPage() {
  const [userType, setUserType] = useState<UserType>('INDIVIDUAL')
  const [step, setStep] = useState(1)
  const [form, setForm] = useState({
    name: '', email: '', password: '', phone: '', address: '',
    companyName: '', department: '', corporateNo: '',
  })

  const update = (k: string, v: string) => setForm(f => ({ ...f, [k]: v }))

  return (
    <div className="min-h-screen bg-gray-50 flex items-center justify-center px-4 py-12">
      <div className="w-full max-w-md">
        <div className="text-center mb-8">
          <h1 className="text-2xl font-bold text-[#1A5276]">アカウント登録</h1>
          <p className="text-gray-500 text-sm mt-1">無料・1分で完了</p>
        </div>

        <div className="bg-white rounded-2xl shadow-md p-8">
          {/* ステップ表示 */}
          <div className="flex items-center gap-2 mb-6">
            {[1, 2].map(s => (
              <div key={s} className="flex items-center gap-2 flex-1">
                <div className={`w-8 h-8 rounded-full flex items-center justify-center text-sm font-bold ${step >= s ? 'bg-[#1A5276] text-white' : 'bg-gray-200 text-gray-400'}`}>
                  {s}
                </div>
                <span className={`text-xs ${step >= s ? 'text-[#1A5276] font-medium' : 'text-gray-400'}`}>
                  {s === 1 ? '会員種別' : '基本情報'}
                </span>
                {s < 2 && <div className={`flex-1 h-0.5 ${step > s ? 'bg-[#1A5276]' : 'bg-gray-200'}`} />}
              </div>
            ))}
          </div>

          {step === 1 && (
            <div>
              <p className="text-sm font-medium text-gray-700 mb-4">ご利用の種別を選択してください</p>
              <div className="grid grid-cols-2 gap-4 mb-6">
                {(['INDIVIDUAL', 'CORPORATE'] as UserType[]).map(type => (
                  <button key={type}
                    onClick={() => setUserType(type)}
                    className={`border-2 rounded-xl p-4 text-center transition-all ${userType === type ? 'border-[#1A5276] bg-blue-50' : 'border-gray-200 hover:border-gray-300'}`}>
                    <div className="text-2xl mb-2">{type === 'INDIVIDUAL' ? '👤' : '🏢'}</div>
                    <p className="font-bold text-sm">{type === 'INDIVIDUAL' ? '個人' : '法人'}</p>
                    <p className="text-xs text-gray-400 mt-1">
                      {type === 'INDIVIDUAL' ? 'LINE/Google連携可' : '与信審査あり'}
                    </p>
                  </button>
                ))}
              </div>
              <button onClick={() => setStep(2)}
                className="w-full bg-[#1A5276] hover:bg-blue-900 text-white font-bold py-3 rounded-xl transition-colors">
                次へ
              </button>
            </div>
          )}

          {step === 2 && (
            <div className="space-y-4">
              <div>
                <label className="text-xs font-medium text-gray-600">お名前 *</label>
                <input className="w-full border border-gray-200 rounded-lg px-3 py-2 mt-1 text-sm focus:outline-none focus:ring-2 focus:ring-[#1A5276]"
                  placeholder="山田 太郎"
                  value={form.name} onChange={e => update('name', e.target.value)} />
              </div>
              <div>
                <label className="text-xs font-medium text-gray-600">メールアドレス *</label>
                <input type="email" className="w-full border border-gray-200 rounded-lg px-3 py-2 mt-1 text-sm focus:outline-none focus:ring-2 focus:ring-[#1A5276]"
                  placeholder="example@email.com"
                  value={form.email} onChange={e => update('email', e.target.value)} />
              </div>
              <div>
                <label className="text-xs font-medium text-gray-600">パスワード *</label>
                <input type="password" className="w-full border border-gray-200 rounded-lg px-3 py-2 mt-1 text-sm focus:outline-none focus:ring-2 focus:ring-[#1A5276]"
                  placeholder="8文字以上"
                  value={form.password} onChange={e => update('password', e.target.value)} />
              </div>
              <div>
                <label className="text-xs font-medium text-gray-600">電話番号</label>
                <input className="w-full border border-gray-200 rounded-lg px-3 py-2 mt-1 text-sm focus:outline-none focus:ring-2 focus:ring-[#1A5276]"
                  placeholder="090-0000-0000"
                  value={form.phone} onChange={e => update('phone', e.target.value)} />
              </div>

              {userType === 'CORPORATE' && (
                <>
                  <div>
                    <label className="text-xs font-medium text-gray-600">会社名 *</label>
                    <input className="w-full border border-gray-200 rounded-lg px-3 py-2 mt-1 text-sm focus:outline-none focus:ring-2 focus:ring-[#1A5276]"
                      placeholder="株式会社〇〇"
                      value={form.companyName} onChange={e => update('companyName', e.target.value)} />
                  </div>
                  <div>
                    <label className="text-xs font-medium text-gray-600">部署名</label>
                    <input className="w-full border border-gray-200 rounded-lg px-3 py-2 mt-1 text-sm focus:outline-none focus:ring-2 focus:ring-[#1A5276]"
                      placeholder="総務部"
                      value={form.department} onChange={e => update('department', e.target.value)} />
                  </div>
                  <div>
                    <label className="text-xs font-medium text-gray-600">法人番号</label>
                    <input className="w-full border border-gray-200 rounded-lg px-3 py-2 mt-1 text-sm focus:outline-none focus:ring-2 focus:ring-[#1A5276]"
                      placeholder="1234567890123"
                      value={form.corporateNo} onChange={e => update('corporateNo', e.target.value)} />
                  </div>
                </>
              )}

              {userType === 'INDIVIDUAL' && (
                <div className="bg-yellow-50 border border-yellow-200 rounded-lg p-3 text-xs text-yellow-800">
                  🪪 登録後、本人確認書類（運転免許証等）のアップロードが必要です。確認完了後にレンタル申込が可能になります。
                </div>
              )}

              <div className="flex gap-3 pt-2">
                <button onClick={() => setStep(1)}
                  className="flex-1 border border-gray-300 text-gray-600 font-medium py-3 rounded-xl text-sm hover:bg-gray-50 transition-colors">
                  戻る
                </button>
                <button className="flex-1 bg-[#1E8B4C] hover:bg-green-600 text-white font-bold py-3 rounded-xl text-sm transition-colors">
                  登録する
                </button>
              </div>
            </div>
          )}

          <p className="text-center text-xs text-gray-400 mt-6">
            すでにアカウントをお持ちの方は{' '}
            <Link href="/auth/login" className="text-[#1A5276] underline">ログイン</Link>
          </p>
        </div>
      </div>
    </div>
  )
}
