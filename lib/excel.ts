// lib/excel.ts
import * as XLSX from 'xlsx'

export type AssetType = 'mutual-funds' | 'etfs' | 'stocks'

export interface Fund {
  id: string
  name: string
  category: string
  sipAmount: number
  startDate: string
  expectedReturn: number
  color: string
  assetType?: 'mutual-funds'
}

export interface ETF {
  id: string
  name: string
  symbol: string
  category: string
  sipAmount: number
  startDate: string
  expectedReturn: number
  color: string
  assetType?: 'etfs'
}

export interface Stock {
  id: string
  name: string
  symbol: string
  category: string
  quantity: number
  buyPrice: number
  startDate: string
  expectedReturn: number
  color: string
  assetType?: 'stocks'
}

export interface Payment {
  id: string
  assetId: string
  assetType: AssetType
  date: string
  amount: number
  quantity?: number
  price?: number
  notes?: string
}

export interface TrackerData {
  funds: Fund[]
  etfs: ETF[]
  stocks: Stock[]
  payments: Payment[]
}

const ASSET_COLORS = ['#C9A84C', '#1A5C3A', '#8B4513', '#4A148C', '#1565C0', '#BF360C', '#00695C', '#6A1B9A', '#E91E63', '#9C27B0']

export function emptyData(): TrackerData {
  return { funds: [], etfs: [], stocks: [], payments: [] }
}

export function parseExcel(file: ArrayBuffer): TrackerData {
  const wb = XLSX.read(file, { type: 'array' })

  const funds: Fund[] = []
  const etfs: ETF[] = []
  const stocks: Stock[] = []
  const payments: Payment[] = []

  // Read Funds sheet
  if (wb.SheetNames.includes('Funds')) {
    const ws = wb.Sheets['Funds']
    const rows = XLSX.utils.sheet_to_json<any>(ws, { defval: '' })
    rows.forEach((r) => {
      if (r.id && r.name) {
        funds.push({
          id: String(r.id),
          name: String(r.name),
          category: String(r.category || 'Equity'),
          sipAmount: Number(r.sipAmount) || 0,
          startDate: String(r.startDate || ''),
          expectedReturn: Number(r.expectedReturn) || 12,
          color: String(r.color || ASSET_COLORS[0]),
          assetType: 'mutual-funds',
        })
      }
    })
  }

  // Read ETFs sheet
  if (wb.SheetNames.includes('ETFs')) {
    const ws = wb.Sheets['ETFs']
    const rows = XLSX.utils.sheet_to_json<any>(ws, { defval: '' })
    rows.forEach((r) => {
      if (r.id && r.name) {
        etfs.push({
          id: String(r.id),
          name: String(r.name),
          symbol: String(r.symbol || ''),
          category: String(r.category || 'Large Cap'),
          sipAmount: Number(r.sipAmount) || 0,
          startDate: String(r.startDate || ''),
          expectedReturn: Number(r.expectedReturn) || 12,
          color: String(r.color || ASSET_COLORS[2]),
          assetType: 'etfs',
        })
      }
    })
  }

  // Read Stocks sheet
  if (wb.SheetNames.includes('Stocks')) {
    const ws = wb.Sheets['Stocks']
    const rows = XLSX.utils.sheet_to_json<any>(ws, { defval: '' })
    rows.forEach((r) => {
      if (r.id && r.name) {
        stocks.push({
          id: String(r.id),
          name: String(r.name),
          symbol: String(r.symbol || ''),
          category: String(r.category || 'Large Cap'),
          quantity: Number(r.quantity) || 0,
          buyPrice: Number(r.buyPrice) || 0,
          startDate: String(r.startDate || ''),
          expectedReturn: Number(r.expectedReturn) || 15,
          color: String(r.color || ASSET_COLORS[4]),
          assetType: 'stocks',
        })
      }
    })
  }

  // Read Payments sheet
  if (wb.SheetNames.includes('Payments')) {
    const ws = wb.Sheets['Payments']
    const rows = XLSX.utils.sheet_to_json<any>(ws, { defval: '' })
    rows.forEach((r) => {
      if (r.id && r.assetId && r.date) {
        payments.push({
          id: String(r.id),
          assetId: String(r.assetId),
          assetType: String(r.assetType) as AssetType,
          date: String(r.date),
          amount: Number(r.amount) || 0,
          quantity: r.quantity ? Number(r.quantity) : undefined,
          price: r.price ? Number(r.price) : undefined,
          notes: r.notes ? String(r.notes) : undefined,
        })
      }
    })
  }

  return { funds, etfs, stocks, payments }
}

export function exportExcel(data: TrackerData): void {
  const wb = XLSX.utils.book_new()

  // Funds sheet
  const fundsData = data.funds.map((f) => ({
    id: f.id,
    name: f.name,
    category: f.category,
    sipAmount: f.sipAmount,
    startDate: f.startDate,
    expectedReturn: f.expectedReturn,
    color: f.color,
  }))
  const fundsWs = XLSX.utils.json_to_sheet(fundsData.length ? fundsData : [
    { id: '', name: '', category: '', sipAmount: '', startDate: '', expectedReturn: '', color: '' }
  ])
  XLSX.utils.book_append_sheet(wb, fundsWs, 'Funds')

  // ETFs sheet
  const etfsData = data.etfs.map((e) => ({
    id: e.id,
    name: e.name,
    symbol: e.symbol,
    category: e.category,
    sipAmount: e.sipAmount,
    startDate: e.startDate,
    expectedReturn: e.expectedReturn,
    color: e.color,
  }))
  const etfsWs = XLSX.utils.json_to_sheet(etfsData.length ? etfsData : [
    { id: '', name: '', symbol: '', category: '', sipAmount: '', startDate: '', expectedReturn: '', color: '' }
  ])
  XLSX.utils.book_append_sheet(wb, etfsWs, 'ETFs')

  // Stocks sheet
  const stocksData = data.stocks.map((s) => ({
    id: s.id,
    name: s.name,
    symbol: s.symbol,
    category: s.category,
    quantity: s.quantity,
    buyPrice: s.buyPrice,
    startDate: s.startDate,
    expectedReturn: s.expectedReturn,
    color: s.color,
  }))
  const stocksWs = XLSX.utils.json_to_sheet(stocksData.length ? stocksData : [
    { id: '', name: '', symbol: '', category: '', quantity: '', buyPrice: '', startDate: '', expectedReturn: '', color: '' }
  ])
  XLSX.utils.book_append_sheet(wb, stocksWs, 'Stocks')

  // Payments sheet
  const paymentsData = data.payments.map((p) => ({
    id: p.id,
    assetId: p.assetId,
    assetType: p.assetType,
    date: p.date,
    amount: p.amount,
    quantity: p.quantity ?? '',
    price: p.price ?? '',
    notes: p.notes ?? '',
  }))
  const paymentsWs = XLSX.utils.json_to_sheet(paymentsData.length ? paymentsData : [
    { id: '', assetId: '', assetType: '', date: '', amount: '', quantity: '', price: '', notes: '' }
  ])
  XLSX.utils.book_append_sheet(wb, paymentsWs, 'Payments')

  // Summary sheet (read-only overview)
  const summaryRows = [
    ...data.funds.map((f) => ({
      'Asset Type': 'Mutual Fund',
      'Name': f.name,
      'Category': f.category,
      'Monthly SIP (₹)': f.sipAmount,
      'Expected Return (%)': f.expectedReturn,
      'Start Date': f.startDate,
    })),
    ...data.etfs.map((e) => ({
      'Asset Type': 'ETF',
      'Name': e.name,
      'Symbol': e.symbol,
      'Category': e.category,
      'Monthly SIP (₹)': e.sipAmount,
      'Expected Return (%)': e.expectedReturn,
      'Start Date': e.startDate,
    })),
    ...data.stocks.map((s) => ({
      'Asset Type': 'Stock',
      'Name': s.name,
      'Symbol': s.symbol,
      'Category': s.category,
      'Quantity': s.quantity,
      'Buy Price (₹)': s.buyPrice,
      'Expected Return (%)': s.expectedReturn,
      'Start Date': s.startDate,
    }))
  ]
  if (summaryRows.length) {
    const summaryWs = XLSX.utils.json_to_sheet(summaryRows)
    XLSX.utils.book_append_sheet(wb, summaryWs, 'Summary')
  }

  XLSX.writeFile(wb, 'investment-tracker.xlsx')
}

// ── Calculation helpers ──────────────────────────────────────────────

export function calcFutureValue(monthlyAmount: number, months: number, annualRate: number): number {
  if (annualRate === 0) return monthlyAmount * months
  const r = annualRate / 100 / 12
  return monthlyAmount * ((Math.pow(1 + r, months) - 1) / r) * (1 + r)
}

export function calcTotalInvested(payments: Payment[], assetId?: string, assetType?: AssetType): number {
  return payments
    .filter((p) => (!assetId || p.assetId === assetId) && (!assetType || p.assetType === assetType))
    .reduce((s, p) => s + p.amount, 0)
}

export function getMonthLabel(ym: string): string {
  const [y, m] = ym.split('-')
  const date = new Date(Number(y), Number(m) - 1)
  return date.toLocaleString('default', { month: 'short', year: '2-digit' })
}

export function generateId(): string {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 6)
}

export function currentYearMonth(): string {
  const d = new Date()
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`
}