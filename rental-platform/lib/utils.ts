export function formatPrice(price: number): string {
  return `¥${price.toLocaleString('ja-JP')}`
}

export function formatDate(dateStr: string): string {
  return new Date(dateStr).toLocaleDateString('ja-JP', {
    year: 'numeric', month: 'long', day: 'numeric'
  })
}

export function calcMinEndDate(startDate: Date, months = 3): Date {
  const d = new Date(startDate)
  d.setMonth(d.getMonth() + months)
  return d
}

export function getStatusLabel(status: string): string {
  const map: Record<string, string> = {
    AVAILABLE: '在庫あり',
    RENTED: '貸出中',
    MAINTENANCE: '整備中',
    RETIRED: '廃棄',
    ACTIVE: '利用中',
    CANCELLED: 'キャンセル',
    COMPLETED: '返却済み',
    PENDING_RETURN: '返却受付中',
    PENDING: '審査中',
    APPROVED: '承認済み',
    REJECTED: '否認',
  }
  return map[status] ?? status
}

export function getGradeLabel(grade: string): string {
  return grade === 'A' ? 'Aランク（良好）' : 'Bランク（外装に軽微な傷あり）'
}
