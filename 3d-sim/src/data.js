export const BOSS = {
  id: 'murapee',
  name: 'むらぴー',
  role: '代表',
  bodyColor: '#003D5C',
  accentColor: '#FFD700',
  desk: [0, 0, 3],
  isLeader: true,
  tasks: ['戦略立案中', '会社を見守り中', '未来を描き中'],
}

export const EMPLOYEES = [
  {
    id: 'lp',
    name: 'LP制作君',
    role: 'LP制作',
    bodyColor: '#0099D4',
    accentColor: '#B0DFF2',
    desk: [-3.5, 0, -1],
    tasks: ['LPデザイン中', 'サイト更新中', 'バナー制作中', 'コーディング中'],
    reportInterval: 10,
  },
  {
    id: 'sales',
    name: '営業企画君',
    role: '営業企画',
    bodyColor: '#27AE60',
    accentColor: '#A9DFB8',
    desk: [3.5, 0, -1],
    tasks: ['提案書作成中', '商談準備中', '顧客リスト整理中', '営業資料更新中'],
    reportInterval: 13,
  },
  {
    id: 'biz',
    name: '経営企画君',
    role: '経営企画',
    bodyColor: '#E67E22',
    accentColor: '#FAD7A0',
    desk: [0, 0, -3.5],
    tasks: ['事業計画作成中', '数値分析中', 'レポート執筆中', '予算策定中'],
    reportInterval: 16,
  },
]

export const ALL_MEMBERS = [BOSS, ...EMPLOYEES]
