export type UserType = 'INDIVIDUAL' | 'CORPORATE'
export type KycStatus = 'PENDING' | 'APPROVED' | 'REJECTED'
export type ProductStatus = 'AVAILABLE' | 'RENTED' | 'MAINTENANCE' | 'RETIRED'
export type ProductGrade = 'A' | 'B'
export type OrderStatus = 'ACTIVE' | 'CANCELLED' | 'COMPLETED' | 'PENDING_RETURN'

export interface Product {
  id: string
  name: string
  brand?: string
  category: string
  grade: ProductGrade
  status: ProductStatus
  monthlyPrice: number
  imageUrl?: string
  description?: string
  sourceLocation?: string
  refurbishedAt?: string
  serialNo?: string
  createdAt: string
}

export interface Order {
  id: string
  userId: string
  productId: string
  status: OrderStatus
  startDate: string
  minEndDate: string
  endDate?: string
  monthlyPrice: number
  deliveryAddress: string
  returnRequestedAt?: string
  product?: Product
  createdAt: string
}

export interface User {
  id: string
  email: string
  name: string
  type: UserType
  kycStatus: KycStatus
  companyName?: string
  department?: string
  phone?: string
  address?: string
  createdAt: string
}
