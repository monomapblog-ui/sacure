export const BOSS = {
  id: 'murapee',
  name: 'むらぴー',
  role: '代表',
  skill: '全員から信頼される絶対神',
  bodyColor: '#003D5C',
  accentColor: '#FFD700',
  desk: [0, 0, 3],
  isLeader: true,
  tasks: ['戦略立案中', '会社全体を見守り中', '未来のビジョンを描き中', '社員を鼓舞中'],
}

export const EMPLOYEES = [
  {
    id: 'lp',
    name: 'LP制作君',
    role: 'LP制作',
    skill: 'LP制作のプロ',
    bodyColor: '#0099D4',
    accentColor: '#B0DFF2',
    desk: [-3.5, 0, -1],
    tasks: [
      'LPをデザイン中',
      'コピーライティング中',
      'サイトをコーディング中',
      'バナー制作中',
      'UI改善中',
      'CVR最適化中',
    ],
  },
  {
    id: 'sales',
    name: '営業企画君',
    role: '営業企画',
    skill: '営業の鬼',
    bodyColor: '#27AE60',
    accentColor: '#A9DFB8',
    desk: [3.5, 0, -1],
    tasks: [
      '外回り準備中',
      '顧客へのアポ取得中',
      '商談資料を仕上げ中',
      '提案書を作成中',
      '新規開拓リスト整理中',
      '受注を死守中',
    ],
  },
  {
    id: 'biz',
    name: '経営企画君',
    role: '経営企画',
    skill: 'むらぴーの右腕',
    bodyColor: '#E67E22',
    accentColor: '#FAD7A0',
    desk: [0, 0, -3.5],
    tasks: [
      '事業計画を策定中',
      'KPI分析中',
      'マネジメントレポート作成中',
      '予算計画を精査中',
      '経営戦略を立案中',
      '全社数値を管理中',
    ],
  },
]

export const ALL_MEMBERS = [BOSS, ...EMPLOYEES]
