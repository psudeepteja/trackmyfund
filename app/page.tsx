"use client"
import { useState, useCallback, useRef, useEffect, useMemo } from "react"
import {
  Upload, Download, Plus, Trash2, TrendingUp, Wallet,
  Calendar, BarChart3, X, Edit3, BarChart2,
  Activity, Search, RefreshCw, Clock
} from "lucide-react"
import {
  AreaChart, Area, XAxis, YAxis, Tooltip, ResponsiveContainer,
  PieChart, Pie, Cell, BarChart, Bar
} from "recharts"
import * as XLSX from "xlsx"

// ─── Types & Constants ────────────────────────────────────────────────────────
const ASSET_COLORS = ['#C9A84C', '#1A5C3A', '#8B4513', '#4A148C', '#1565C0', '#BF360C', '#00695C', '#6A1B9A', '#E91E63', '#9C27B0']
const CATEGORIES = ['Equity', 'Debt', 'Hybrid', 'ELSS', 'Index', 'Gold', 'International', 'Other']
const ETF_CATEGORIES = ['Large Cap', 'Mid Cap', 'Small Cap', 'Sectoral', 'Gold', 'International', 'Debt', 'Other']
const STOCK_CATEGORIES = ['Large Cap', 'Mid Cap', 'Small Cap', 'Dividend', 'Growth', 'PSU', 'Banking', 'IT', 'Pharma', 'Other']
const STORAGE_KEY = 'investment-tracker-data'

// Base Asset Interface
interface BaseAsset {
  id: string
  name: string
  category: string
  startDate: string
  expectedReturn: number
  color: string
  assetType: string
}

interface MutualFund extends BaseAsset {
  sipAmount: number
  assetType: 'mutual-funds'
}

interface ETF extends BaseAsset {
  symbol: string
  sipAmount: number
  assetType: 'etfs'
}

interface Stock extends BaseAsset {
  symbol: string
  quantity: number
  buyPrice: number
  assetType: 'stocks'
}

interface Payment {
  id: string
  assetId: string
  assetType: string
  date: string
  amount: number
  quantity?: number
  price?: number
  notes?: string
  nav?: number
  units?: number
}

interface AppData {
  funds: MutualFund[]
  etfs: ETF[]
  stocks: Stock[]
  payments: Payment[]
}

// Type guards
function isMutualFund(asset: any): asset is MutualFund {
  return asset && 'sipAmount' in asset && asset.assetType === 'mutual-funds'
}

function isETF(asset: any): asset is ETF {
  return asset && 'symbol' in asset && 'sipAmount' in asset && asset.assetType === 'etfs'
}

function isStock(asset: any): asset is Stock {
  return asset && 'symbol' in asset && 'quantity' in asset && 'buyPrice' in asset && asset.assetType === 'stocks'
}

function generateId(): string {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 6)
}

function currentYearMonth(): string {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`
}

function getMonthLabel(ym: string): string {
  const [y, m] = ym.split('-');
  return new Date(+y, +m - 1).toLocaleString('default', { month: 'short', year: '2-digit' })
}

function calcFutureValue(monthly: number, months: number, rate: number): number {
  if (!rate) return monthly * months;
  const r = rate / 100 / 12;
  return monthly * ((Math.pow(1 + r, months) - 1) / r) * (1 + r)
}

function calcTotalInvested(payments: Payment[], assetId?: string, assetType?: string): number {
  return payments.filter(p => (!assetId || p.assetId === assetId) && (!assetType || p.assetType === assetType)).reduce((s, p) => s + p.amount, 0)
}

function emptyData(): AppData {
  return { funds: [], etfs: [], stocks: [], payments: [] }
}

// ─── Excel ────────────────────────────────────────────────────────────────────
function parseExcel(buf: ArrayBuffer): AppData {
  const wb = XLSX.read(buf, { type: 'array' });
  const funds: MutualFund[] = [], etfs: ETF[] = [], stocks: Stock[] = [], payments: Payment[] = []
  if (wb.SheetNames.includes('Funds')) {
    XLSX.utils.sheet_to_json(wb.Sheets['Funds'], { defval: '' }).forEach((r: any) => {
      if (r.id && r.name) funds.push({
        id: String(r.id),
        name: String(r.name),
        category: String(r.category || 'Equity'),
        sipAmount: +r.sipAmount || 0,
        startDate: String(r.startDate || ''),
        expectedReturn: +r.expectedReturn || 12,
        color: String(r.color || ASSET_COLORS[0]),
        assetType: 'mutual-funds'
      })
    })
  }
  if (wb.SheetNames.includes('ETFs')) {
    XLSX.utils.sheet_to_json(wb.Sheets['ETFs'], { defval: '' }).forEach((r: any) => {
      if (r.id && r.name) etfs.push({
        id: String(r.id),
        name: String(r.name),
        symbol: String(r.symbol || ''),
        category: String(r.category || 'Large Cap'),
        sipAmount: +r.sipAmount || 0,
        startDate: String(r.startDate || ''),
        expectedReturn: +r.expectedReturn || 12,
        color: String(r.color || ASSET_COLORS[2]),
        assetType: 'etfs'
      })
    })
  }
  if (wb.SheetNames.includes('Stocks')) {
    XLSX.utils.sheet_to_json(wb.Sheets['Stocks'], { defval: '' }).forEach((r: any) => {
      if (r.id && r.name) stocks.push({
        id: String(r.id),
        name: String(r.name),
        symbol: String(r.symbol || ''),
        category: String(r.category || 'Large Cap'),
        quantity: +r.quantity || 0,
        buyPrice: +r.buyPrice || 0,
        startDate: String(r.startDate || ''),
        expectedReturn: +r.expectedReturn || 15,
        color: String(r.color || ASSET_COLORS[4]),
        assetType: 'stocks'
      })
    })
  }
  if (wb.SheetNames.includes('Payments')) {
    XLSX.utils.sheet_to_json(wb.Sheets['Payments'], { defval: '' }).forEach((r: any) => {
      if (r.id && r.assetId && r.date) payments.push({
        id: String(r.id),
        assetId: String(r.assetId),
        assetType: String(r.assetType),
        date: String(r.date),
        amount: +r.amount || 0,
        quantity: r.quantity ? +r.quantity : undefined,
        price: r.price ? +r.price : undefined,
        notes: r.notes ? String(r.notes) : undefined,
        nav: r.nav ? +r.nav : undefined,
        units: r.units ? +r.units : undefined
      })
    })
  }
  return { funds, etfs, stocks, payments }
}

function exportExcel(data: AppData): void {
  const wb = XLSX.utils.book_new()
  const addSheet = (rows: any[], name: string) => XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows.length ? rows : [{}]), name)
  addSheet(data.funds.map(f => ({ id: f.id, name: f.name, category: f.category, sipAmount: f.sipAmount, startDate: f.startDate, expectedReturn: f.expectedReturn, color: f.color })), 'Funds')
  addSheet(data.etfs.map(e => ({ id: e.id, name: e.name, symbol: e.symbol, category: e.category, sipAmount: e.sipAmount, startDate: e.startDate, expectedReturn: e.expectedReturn, color: e.color })), 'ETFs')
  addSheet(data.stocks.map(s => ({ id: s.id, name: s.name, symbol: s.symbol, category: s.category, quantity: s.quantity, buyPrice: s.buyPrice, startDate: s.startDate, expectedReturn: s.expectedReturn, color: s.color })), 'Stocks')
  addSheet(data.payments.map(p => ({ id: p.id, assetId: p.assetId, assetType: p.assetType, date: p.date, amount: p.amount, quantity: p.quantity ?? '', price: p.price ?? '', notes: p.notes ?? '', nav: p.nav ?? '', units: p.units ?? '' })), 'Payments')
  XLSX.writeFile(wb, 'investment-tracker.xlsx')
}

// ─── Component Props Types ────────────────────────────────────────────────────
interface ModalProps {
  title: string
  onClose: () => void
  children: React.ReactNode
  wide?: boolean
}

interface FieldProps {
  label: string
  children: React.ReactNode
}

interface ColorPickerProps {
  value: string
  onChange: (color: string) => void
}

interface FundModalProps {
  fund?: MutualFund
  onSave: (fund: MutualFund) => void
  onClose: () => void
}

interface ETFModalProps {
  etf?: ETF
  onSave: (etf: ETF) => void
  onClose: () => void
}

interface StockModalProps {
  stock?: Stock
  onSave: (stock: Stock) => void
  onClose: () => void
}

interface InvestmentModalProps {
  assetType: string
  funds: MutualFund[]
  etfs: ETF[]
  stocks: Stock[]
  payment?: Payment
  preSelectedAssetId?: string | null
  onSave: (payment: Payment) => void
  onClose: () => void
}

interface ProjectionModalProps {
  funds: MutualFund[]
  etfs: ETF[]
  stocks: Stock[]
  onClose: () => void
}

interface AssetHistoryModalProps {
  asset: MutualFund | ETF | Stock
  payments: Payment[]
  onClose: () => void
  onEdit: (payment: Payment) => void
  onDelete: (id: string) => void
}

interface SearchBarProps {
  value: string
  onChange: (value: string) => void
  placeholder: string
}

// ─── Modal Shell ──────────────────────────────────────────────────────────────
function Modal({ title, onClose, children, wide = false }: ModalProps) {
  return (
    <div className="fixed inset-0 z-50 flex items-end sm:items-center justify-center p-0 sm:p-4" onClick={onClose}>
      <div className="absolute inset-0 bg-black/40 backdrop-blur-sm" />
      <div
        className={`relative bg-[#fff] rounded-t-3xl sm:rounded-2xl shadow-2xl border border-[#C9A84C]/20 w-full ${wide ? 'sm:max-w-2xl' : 'sm:max-w-md'} p-5 sm:p-6 max-h-[92vh] overflow-y-auto`}
        onClick={e => e.stopPropagation()}>
        <div className="sm:hidden w-10 h-1 bg-[#C9A84C]/30 rounded-full mx-auto mb-4" />
        <div className="flex items-center justify-between mb-5">
          <h3 className="text-base font-semibold text-[#0D0D0D]">{title}</h3>
          <button onClick={onClose} className="text-[#8A8070] hover:text-[#0D0D0D] transition-colors p-1"><X size={16} /></button>
        </div>
        {children}
      </div>
    </div>
  )
}

// ─── Asset History Modal ──────────────────────────────────────────────────────
function AssetHistoryModal({ asset, payments, onClose, onEdit, onDelete }: AssetHistoryModalProps) {
  const isMF = asset.assetType === 'mutual-funds'
  const isETFAsset = asset.assetType === 'etfs'
  const isStockAsset = asset.assetType === 'stocks'

  const sorted = [...payments].sort((a, b) => b.date.localeCompare(a.date))
  const totalInvested = payments.reduce((s, p) => s + p.amount, 0)
  const totalUnits = payments.reduce((s, p) => s + (p.units || 0), 0)
  const totalQty = payments.reduce((s, p) => s + (p.quantity || 0), 0)
  const avgNav = isMF && totalUnits > 0 ? totalInvested / totalUnits : null
  const avgPrice = (isETFAsset || isStockAsset) && totalQty > 0 ? totalInvested / totalQty : null

  return (
    <Modal title={`${asset.name} — History`} onClose={onClose} wide>
      {/* Summary strip */}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 mb-5">
        <div className="bg-[#F5F0E8] rounded-xl p-3">
          <p className="text-xs text-[#8A8070]">Total Invested</p>
          <p className="font-bold text-[#1A5C3A] text-sm">₹{totalInvested.toLocaleString('en-IN')}</p>
        </div>
        <div className="bg-[#F5F0E8] rounded-xl p-3">
          <p className="text-xs text-[#8A8070]">Entries</p>
          <p className="font-bold text-[#0D0D0D] text-sm">{payments.length}</p>
        </div>
        {isMF && totalUnits > 0 && (
          <>
            <div className="bg-[#F5F0E8] rounded-xl p-3">
              <p className="text-xs text-[#8A8070]">Total Units</p>
              <p className="font-bold text-[#4A148C] text-sm">{totalUnits.toFixed(3)}</p>
            </div>
            <div className="bg-[#F5F0E8] rounded-xl p-3">
              <p className="text-xs text-[#8A8070]">Avg NAV</p>
              <p className="font-bold text-[#8B6914] text-sm">₹{avgNav?.toFixed(3)}</p>
            </div>
          </>
        )}
        {(isETFAsset || isStockAsset) && totalQty > 0 && (
          <>
            <div className="bg-[#F5F0E8] rounded-xl p-3">
              <p className="text-xs text-[#8A8070]">Total Qty</p>
              <p className="font-bold text-[#4A148C] text-sm">{totalQty}</p>
            </div>
            {avgPrice !== null && (
              <div className="bg-[#F5F0E8] rounded-xl p-3">
                <p className="text-xs text-[#8A8070]">Avg Price</p>
                <p className="font-bold text-[#8B6914] text-sm">₹{avgPrice.toFixed(2)}</p>
              </div>
            )}
          </>
        )}
      </div>

      {payments.length === 0 ? (
        <div className="text-center py-10">
          <Calendar size={32} className="text-[#C9A84C]/40 mx-auto mb-2" />
          <p className="text-sm text-[#8A8070]">No investments logged yet for this asset</p>
        </div>
      ) : (
        <>
          {/* Mobile list */}
          <div className="sm:hidden space-y-2">
            {sorted.map(p => (
              <div key={p.id} className="border border-[#C9A84C]/15 rounded-xl p-3">
                <div className="flex items-center justify-between gap-2">
                  <div>
                    <p className="text-sm font-medium text-[#0D0D0D]">{getMonthLabel(p.date)}</p>
                    <p className="text-xs text-[#8A8070]">{p.date}</p>
                  </div>
                  <p className="font-semibold text-[#1A5C3A]">₹{p.amount.toLocaleString('en-IN')}</p>
                </div>
                {(p.nav || p.units || p.quantity || p.price) && (
                  <div className="flex flex-wrap gap-1.5 mt-2">
                    {p.nav && <span className="text-xs bg-[#F5F0E8] text-[#8A8070] px-2 py-0.5 rounded-lg">NAV ₹{p.nav}</span>}
                    {p.units && <span className="text-xs bg-[#F5F0E8] text-[#4A148C] px-2 py-0.5 rounded-lg">{p.units} units</span>}
                    {p.price && !p.nav && <span className="text-xs bg-[#F5F0E8] text-[#8A8070] px-2 py-0.5 rounded-lg">₹{p.price}/unit</span>}
                    {p.quantity && <span className="text-xs bg-[#F5F0E8] text-[#0D0D0D] px-2 py-0.5 rounded-lg">Qty {p.quantity}</span>}
                  </div>
                )}
                <div className="flex items-center justify-between mt-2">
                  <span className="text-xs text-[#8A8070] truncate">{p.notes || ''}</span>
                  <div className="flex gap-1">
                    <button onClick={() => onEdit(p)} className="p-1.5 text-[#8A8070] hover:text-[#0D0D0D] rounded-lg"><Edit3 size={12} /></button>
                    <button onClick={() => onDelete(p.id)} className="p-1.5 text-[#8A8070] hover:text-red-500 rounded-lg"><Trash2 size={12} /></button>
                  </div>
                </div>
              </div>
            ))}
          </div>

          {/* Desktop table */}
          <div className="hidden sm:block border border-[#C9A84C]/15 rounded-xl overflow-hidden">
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="border-b border-[#C9A84C]/15 bg-[#F5F0E8]/60">
                    <th className="text-left px-4 py-3 text-xs font-semibold text-[#4a4a4a] tracking-wide">Month</th>
                    <th className="text-right px-4 py-3 text-xs font-semibold text-[#4a4a4a] tracking-wide">Amount</th>
                    {isMF && <>
                      <th className="text-right px-4 py-3 text-xs font-semibold text-[#4a4a4a] tracking-wide">NAV</th>
                      <th className="text-right px-4 py-3 text-xs font-semibold text-[#4a4a4a] tracking-wide">Units</th>
                    </>}
                    {(isETFAsset || isStockAsset) && <>
                      <th className="text-right px-4 py-3 text-xs font-semibold text-[#4a4a4a] tracking-wide">Price</th>
                      <th className="text-right px-4 py-3 text-xs font-semibold text-[#4a4a4a] tracking-wide">Qty</th>
                    </>}
                    <th className="text-left px-4 py-3 text-xs font-semibold text-[#4a4a4a] tracking-wide">Notes</th>
                    <th className="px-4 py-3" />
                  </tr>
                </thead>
                <tbody>
                  {sorted.map(p => (
                    <tr key={p.id} className="border-b border-[#C9A84C]/10 hover:bg-[#F5F0E8]/40 transition-colors">
                      <td className="px-4 py-3 text-sm text-[#4a4a4a]">{getMonthLabel(p.date)}</td>
                      <td className="px-4 py-3 text-right font-medium text-[#1A5C3A]">₹{p.amount.toLocaleString('en-IN')}</td>
                      {isMF && <>
                        <td className="px-4 py-3 text-right text-[#4a4a4a]">{p.nav ? `₹${p.nav}` : '—'}</td>
                        <td className="px-4 py-3 text-right text-[#4A148C] font-medium">{p.units || '—'}</td>
                      </>}
                      {(isETFAsset || isStockAsset) && <>
                        <td className="px-4 py-3 text-right text-[#4a4a4a]">{p.price ? `₹${p.price}` : '—'}</td>
                        <td className="px-4 py-3 text-right text-[#4a4a4a]">{p.quantity || '—'}</td>
                      </>}
                      <td className="px-4 py-3 text-[#4a4a4a] text-xs truncate max-w-[100px]">{p.notes || '—'}</td>
                      <td className="px-4 py-3">
                        <div className="flex gap-1 justify-end">
                          <button onClick={() => onEdit(p)} className="p-1.5 text-[#8A8070] hover:text-[#0D0D0D] rounded-lg hover:bg-[#F5F0E8] transition-all"><Edit3 size={13} /></button>
                          <button onClick={() => onDelete(p.id)} className="p-1.5 text-[#8A8070] hover:text-red-500 rounded-lg hover:bg-[#F5F0E8] transition-all"><Trash2 size={13} /></button>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
                <tfoot>
                  <tr className="bg-[#F5F0E8]/30">
                    <td className="px-4 py-3 text-xs font-medium text-[#8A8070]">Total ({payments.length} entries)</td>
                    <td className="px-4 py-3 text-right font-bold text-[#1A5C3A]">₹{totalInvested.toLocaleString('en-IN')}</td>
                    {isMF && <>
                      <td className="px-4 py-3 text-right text-xs text-[#8A8070]">{avgNav ? `Avg ₹${avgNav.toFixed(3)}` : '—'}</td>
                      <td className="px-4 py-3 text-right font-bold text-[#4A148C]">{totalUnits > 0 ? totalUnits.toFixed(3) : '—'}</td>
                    </>}
                    {(isETFAsset || isStockAsset) && <>
                      <td className="px-4 py-3 text-right text-xs text-[#8A8070]">{avgPrice ? `Avg ₹${avgPrice.toFixed(2)}` : '—'}</td>
                      <td className="px-4 py-3 text-right font-bold text-[#4A148C]">{totalQty > 0 ? totalQty : '—'}</td>
                    </>}
                    <td colSpan={99} />
                  </tr>
                </tfoot>
              </table>
            </div>
          </div>
        </>
      )}
    </Modal>
  )
}

// ─── Fund Modal ───────────────────────────────────────────────────────────────
function FundModal({ fund, onSave, onClose }: FundModalProps) {
  const [form, setForm] = useState({
    name: fund?.name ?? '',
    category: fund?.category ?? 'Equity',
    sipAmount: fund?.sipAmount ?? 1000,
    startDate: fund?.startDate ?? currentYearMonth(),
    expectedReturn: fund?.expectedReturn ?? 12,
    color: fund?.color ?? ASSET_COLORS[0]
  })
  const setField = (k: keyof typeof form, v: any) => setForm(p => ({ ...p, [k]: v }))
  return (
    <Modal title={fund ? 'Edit Mutual Fund' : 'Add Mutual Fund'} onClose={onClose}>
      <div className="space-y-4">
        <Field label="Fund Name"><input className="inp" placeholder="e.g. Mirae Asset Large Cap" value={form.name} onChange={e => setField('name', e.target.value)} /></Field>
        <div className="grid grid-cols-2 gap-3">
          <Field label="Category"><select className="inp" value={form.category} onChange={e => setField('category', e.target.value)}>{CATEGORIES.map(c => <option key={c}>{c}</option>)}</select></Field>
          <Field label="Monthly SIP (₹)"><input className="inp" type="number" value={form.sipAmount} onChange={e => setField('sipAmount', +e.target.value)} /></Field>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <Field label="Start Month"><input className="inp" type="month" value={form.startDate} onChange={e => setField('startDate', e.target.value)} /></Field>
          <Field label="Expected CAGR %"><input className="inp" type="number" value={form.expectedReturn} onChange={e => setField('expectedReturn', +e.target.value)} /></Field>
        </div>
        <ColorPicker value={form.color} onChange={c => setField('color', c)} />
        <div className="flex gap-3 pt-2">
          <button className="btn-outline flex-1" onClick={onClose}>Cancel</button>
          <button className="btn-primary flex-1" onClick={() => { if (!form.name.trim()) return; onSave({ ...form, id: fund?.id ?? generateId(), assetType: 'mutual-funds' as const }) }}>
            {fund ? 'Update' : 'Add Fund'}
          </button>
        </div>
      </div>
    </Modal>
  )
}

// ─── ETF Modal ────────────────────────────────────────────────────────────────
function ETFModal({ etf, onSave, onClose }: ETFModalProps) {
  const [form, setForm] = useState({
    name: etf?.name ?? '',
    symbol: etf?.symbol ?? '',
    category: etf?.category ?? 'Large Cap',
    sipAmount: etf?.sipAmount ?? 1000,
    startDate: etf?.startDate ?? currentYearMonth(),
    expectedReturn: etf?.expectedReturn ?? 12,
    color: etf?.color ?? ASSET_COLORS[2]
  })
  const setField = (k: keyof typeof form, v: any) => setForm(p => ({ ...p, [k]: v }))
  return (
    <Modal title={etf ? 'Edit ETF' : 'Add ETF'} onClose={onClose}>
      <div className="space-y-4">
        <Field label="ETF Name"><input className="inp" placeholder="e.g. Nippon ETF Nifty50" value={form.name} onChange={e => setField('name', e.target.value)} /></Field>
        <Field label="Symbol"><input className="inp" placeholder="e.g. NIPPONETF" value={form.symbol} onChange={e => setField('symbol', e.target.value.toUpperCase())} /></Field>
        <div className="grid grid-cols-2 gap-3">
          <Field label="Category"><select className="inp" value={form.category} onChange={e => setField('category', e.target.value)}>{ETF_CATEGORIES.map(c => <option key={c}>{c}</option>)}</select></Field>
          <Field label="Monthly SIP (₹)"><input className="inp" type="number" value={form.sipAmount} onChange={e => setField('sipAmount', +e.target.value)} /></Field>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <Field label="Start Month"><input className="inp" type="month" value={form.startDate} onChange={e => setField('startDate', e.target.value)} /></Field>
          <Field label="Expected CAGR %"><input className="inp" type="number" value={form.expectedReturn} onChange={e => setField('expectedReturn', +e.target.value)} /></Field>
        </div>
        <ColorPicker value={form.color} onChange={c => setField('color', c)} />
        <div className="flex gap-3 pt-2">
          <button className="btn-outline flex-1" onClick={onClose}>Cancel</button>
          <button className="btn-primary flex-1" onClick={() => { if (!form.name.trim() || !form.symbol.trim()) return; onSave({ ...form, id: etf?.id ?? generateId(), assetType: 'etfs' as const }) }}>
            {etf ? 'Update' : 'Add ETF'}
          </button>
        </div>
      </div>
    </Modal>
  )
}

// ─── Stock Modal ──────────────────────────────────────────────────────────────
function StockModal({ stock, onSave, onClose }: StockModalProps) {
  const [form, setForm] = useState({
    name: stock?.name ?? '',
    symbol: stock?.symbol ?? '',
    category: stock?.category ?? 'Large Cap',
    quantity: stock?.quantity ?? 0,
    buyPrice: stock?.buyPrice ?? 0,
    startDate: stock?.startDate ?? currentYearMonth(),
    expectedReturn: stock?.expectedReturn ?? 15,
    color: stock?.color ?? ASSET_COLORS[4]
  })
  const setField = (k: keyof typeof form, v: any) => setForm(p => ({ ...p, [k]: v }))
  return (
    <Modal title={stock ? 'Edit Stock' : 'Add Stock'} onClose={onClose}>
      <div className="space-y-4">
        <Field label="Company Name"><input className="inp" placeholder="e.g. Reliance Industries" value={form.name} onChange={e => setField('name', e.target.value)} /></Field>
        <Field label="Symbol"><input className="inp" placeholder="e.g. RELIANCE" value={form.symbol} onChange={e => setField('symbol', e.target.value.toUpperCase())} /></Field>
        <div className="grid grid-cols-2 gap-3">
          <Field label="Category"><select className="inp" value={form.category} onChange={e => setField('category', e.target.value)}>{STOCK_CATEGORIES.map(c => <option key={c}>{c}</option>)}</select></Field>
          <Field label="Quantity"><input className="inp" type="number" value={form.quantity} onChange={e => setField('quantity', +e.target.value)} /></Field>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <Field label="Buy Price (₹)"><input className="inp" type="number" value={form.buyPrice} onChange={e => setField('buyPrice', +e.target.value)} /></Field>
          <Field label="Start Month"><input className="inp" type="month" value={form.startDate} onChange={e => setField('startDate', e.target.value)} /></Field>
        </div>
        <Field label="Expected CAGR %"><input className="inp" type="number" value={form.expectedReturn} onChange={e => setField('expectedReturn', +e.target.value)} /></Field>
        <ColorPicker value={form.color} onChange={c => setField('color', c)} />
        <div className="flex gap-3 pt-2">
          <button className="btn-outline flex-1" onClick={onClose}>Cancel</button>
          <button className="btn-primary flex-1" onClick={() => { if (!form.name.trim() || !form.symbol.trim()) return; onSave({ ...form, id: stock?.id ?? generateId(), assetType: 'stocks' as const }) }}>
            {stock ? 'Update' : 'Add Stock'}
          </button>
        </div>
      </div>
    </Modal>
  )
}

// ─── Investment Modal ─────────────────────────────────────────────────────────
function InvestmentModal({ assetType, funds, etfs, stocks, payment, preSelectedAssetId, onSave, onClose }: InvestmentModalProps) {
  const getAssets = (): (MutualFund | ETF | Stock)[] => {
    if (assetType === 'mutual-funds') return funds
    if (assetType === 'etfs') return etfs
    return stocks
  }
  const assets = getAssets()

  const defaultAssetId = preSelectedAssetId || payment?.assetId || (assets[0]?.id ?? '')
  const lockedAsset = preSelectedAssetId ? assets.find(a => a.id === preSelectedAssetId) : null

  const getDefaultAmount = () => {
    if (payment?.amount) return payment.amount
    const defaultAsset = assets.find(a => a.id === defaultAssetId)
    if (defaultAsset && ('sipAmount' in defaultAsset)) {
      return defaultAsset.sipAmount
    }
    return 1000
  }

  const [form, setForm] = useState({
    assetId: defaultAssetId,
    date: payment?.date ?? currentYearMonth(),
    amount: getDefaultAmount(),
    quantity: payment?.quantity?.toString() ?? '',
    price: payment?.price?.toString() ?? '',
    nav: payment?.nav?.toString() ?? '',
    units: payment?.units?.toString() ?? '',
    notes: payment?.notes ?? '',
  })

  const setField = (k: keyof typeof form, v: any) => setForm(p => ({ ...p, [k]: v }))

  const isMF = assetType === 'mutual-funds'
  const isStock = assetType === 'stocks'
  const isETF = assetType === 'etfs'

  const handleNavChange = (v: string) => {
    setField('nav', v)
    if (v && form.amount) setField('units', (+form.amount / +v).toFixed(3))
  }
  const handleAmountChange = (v: string) => {
    setField('amount', v)
    if (form.nav && v) setField('units', (+v / +form.nav).toFixed(3))
  }
  const handleQtyChange = (v: string) => {
    setField('quantity', v)
    if (form.price && v) setField('amount', (+v * +form.price).toFixed(2))
  }
  const handlePriceChange = (v: string) => {
    setField('price', v)
    if (form.quantity && v) setField('amount', (+form.quantity * +v).toFixed(2))
  }

  const handleSave = () => {
    if (!form.assetId || !form.date) return
    const p: Payment = {
      id: payment?.id ?? generateId(),
      assetId: form.assetId,
      assetType,
      date: form.date,
      amount: +form.amount
    }
    if (form.quantity) p.quantity = +form.quantity
    if (form.price) p.price = +form.price
    if (form.nav) p.nav = +form.nav
    if (form.units) p.units = +form.units
    if (form.notes) p.notes = String(form.notes)
    onSave(p)
  }

  const assetLabel = isMF ? 'Fund' : isETF ? 'ETF' : 'Stock'

  return (
    <Modal title={payment ? 'Edit Investment' : `Log ${assetLabel} Investment`} onClose={onClose}>
      <div className="space-y-4">
        <Field label={assetLabel}>
          {lockedAsset ? (
            <div className="inp flex items-center gap-2 cursor-not-allowed bg-[#fff]">
              <div className="w-2.5 h-2.5 rounded-full shrink-0" style={{ background: lockedAsset.color }} />
              <span className="truncate font-medium text-[#0D0D0D]">{lockedAsset.name}</span>
              {'symbol' in lockedAsset && <span className="text-[#8A8070] text-xs shrink-0">({lockedAsset.symbol})</span>}
            </div>
          ) : (
            <select className="inp" value={form.assetId} onChange={e => setField('assetId', e.target.value)}>
              {assets.map(a => <option key={a.id} value={a.id}>{a.name}{'symbol' in a ? ` (${a.symbol})` : ''}</option>)}
            </select>
          )}
        </Field>

        <div className="grid grid-cols-2 gap-3">
          <Field label="Month"><input className="inp" type="month" value={form.date} onChange={e => setField('date', e.target.value)} /></Field>
          <Field label="Amount (₹)"><input className="inp" type="number" value={form.amount} onChange={e => handleAmountChange(e.target.value)} /></Field>
        </div>

        {isMF && (
          <div className="grid grid-cols-2 gap-3">
            <Field label="NAV (₹)">
              <input className="inp" type="number" step="0.001" placeholder="e.g. 45.320" value={form.nav} onChange={e => handleNavChange(e.target.value)} />
            </Field>
            <Field label="Units Allotted">
              <input className="inp" type="number" step="0.001" placeholder="Auto-calc" value={form.units} onChange={e => setField('units', e.target.value)} />
            </Field>
          </div>
        )}

        {isETF && (
          <div className="grid grid-cols-2 gap-3">
            <Field label="Market Price (₹)">
              <input className="inp" type="number" step="0.01" placeholder="e.g. 230.50" value={form.price} onChange={e => handlePriceChange(e.target.value)} />
            </Field>
            <Field label="Units / Qty">
              <input className="inp" type="number" step="1" placeholder="e.g. 5" value={form.quantity} onChange={e => handleQtyChange(e.target.value)} />
            </Field>
          </div>
        )}

        {isStock && (
          <div className="grid grid-cols-2 gap-3">
            <Field label="Buy Price (₹)">
              <input className="inp" type="number" step="0.01" placeholder="e.g. 2450.50" value={form.price} onChange={e => handlePriceChange(e.target.value)} />
            </Field>
            <Field label="Quantity">
              <input className="inp" type="number" step="1" placeholder="e.g. 10" value={form.quantity} onChange={e => handleQtyChange(e.target.value)} />
            </Field>
          </div>
        )}

        {isMF && form.nav && form.amount && (
          <p className="text-xs text-[#1A5C3A] bg-[#1A5C3A]/8 rounded-lg px-3 py-1.5">
            📊 {(+form.amount / +form.nav).toFixed(3)} units @ ₹{form.nav} NAV
          </p>
        )}
        {(isETF || isStock) && form.price && form.quantity && (
          <p className="text-xs text-[#1A5C3A] bg-[#1A5C3A]/8 rounded-lg px-3 py-1.5">
            📊 {form.quantity} units × ₹{form.price} = ₹{(+form.quantity * +form.price).toFixed(2)}
          </p>
        )}

        <Field label="Notes"><input className="inp" placeholder="optional" value={form.notes} onChange={e => setField('notes', e.target.value)} /></Field>
        <div className="flex gap-3 pt-2">
          <button className="btn-outline flex-1" onClick={onClose}>Cancel</button>
          <button className="btn-primary flex-1" onClick={handleSave}>{payment ? 'Update' : 'Log Investment'}</button>
        </div>
      </div>
    </Modal>
  )
}

// ─── Projection Modal ─────────────────────────────────────────────────────────
function ProjectionModal({ funds, etfs, stocks, onClose }: ProjectionModalProps) {
  const [years, setYears] = useState(10)
  const allAssets = [...funds, ...etfs, ...stocks]
  const getMonthly = (a: MutualFund | ETF | Stock): number => {
    if (isMutualFund(a) || isETF(a)) return a.sipAmount
    if (isStock(a)) return (a.quantity * a.buyPrice) / 12
    return 0
  }
  const data = Array.from({ length: years }, (_, i) => {
    const month = (i + 1) * 12
    let ti = 0, tv = 0
    allAssets.forEach(a => {
      const m = getMonthly(a)
      ti += m * month
      tv += calcFutureValue(m, month, a.expectedReturn)
    })
    return {
      year: `Yr ${i + 1}`,
      'Total Invested': Math.round(ti),
      'Projected Value': Math.round(tv)
    }
  })
  const total = allAssets.reduce((s, a) => s + getMonthly(a), 0)
  const proj = data[years - 1]?.['Projected Value'] ?? 0
  return (
    <Modal title="Portfolio Projection" onClose={onClose} wide>
      <div className="flex flex-wrap gap-2 mb-4">
        <span className="text-sm text-[#8A8070] self-center">Project for</span>
        {[3, 5, 10, 15, 20, 25, 30].map(y => (
          <button key={y} onClick={() => setYears(y)} className={`px-3 py-1.5 rounded-lg text-sm font-medium transition-all ${years === y ? 'bg-[#C9A84C] text-[#0D0D0D]' : 'bg-[#fff] text-[#8A8070] hover:text-[#0D0D0D]'}`}>{y}Y</button>
        ))}
      </div>
      <div className="grid grid-cols-3 gap-3 mb-4">
        {[
          ['Monthly SIP', '₹' + total.toLocaleString('en-IN'), 'text-[#1A5C3A]'],
          [`Total Invested (${years}Y)`, '₹' + (total * years * 12).toLocaleString('en-IN'), 'text-[#0D0D0D]'],
          ['Projected Value', '₹' + Math.round(proj / 100000).toLocaleString('en-IN') + 'L', 'text-[#8B6914]']
        ].map(([l, v, c]) => (
          <div key={l} className="bg-[#fff] rounded-xl p-3">
            <p className="text-xs text-[#8A8070]">{l}</p>
            <p className={`font-bold text-lg ${c}`}>{v}</p>
          </div>
        ))}
      </div>
      <ResponsiveContainer width="100%" height={200}>
        <AreaChart data={data}>
          <defs>
            <linearGradient id="gI" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="#8B6914" stopOpacity={0.3} /><stop offset="95%" stopColor="#8B6914" stopOpacity={0} /></linearGradient>
            <linearGradient id="gV" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="#1A5C3A" stopOpacity={0.4} /><stop offset="95%" stopColor="#1A5C3A" stopOpacity={0} /></linearGradient>
          </defs>
          <XAxis dataKey="year" tick={{ fontSize: 11, fill: '#8A8070' }} />
          <YAxis tickFormatter={v => `₹${(v / 100000).toFixed(0)}L`} tick={{ fontSize: 11, fill: '#8A8070' }} />
          <Tooltip formatter={(v: number) => [`₹${Number(v).toLocaleString('en-IN')}`, '']} contentStyle={{ background: '#FDF8F0', border: '1px solid #C9A84C33', borderRadius: 12 }} />
          <Area type="monotone" dataKey="Total Invested" stroke="#8B6914" fill="url(#gI)" strokeWidth={2} />
          <Area type="monotone" dataKey="Projected Value" stroke="#1A5C3A" fill="url(#gV)" strokeWidth={2} />
        </AreaChart>
      </ResponsiveContainer>
      <p className="text-xs text-[#8A8070] text-center mt-2">Estimates based on expected CAGR. Actual returns may vary.</p>
    </Modal>
  )
}

// ─── Small helpers ────────────────────────────────────────────────────────────
function Field({ label, children }: FieldProps) {
  return <div><label className="text-xs font-medium text-[#8A8070] mb-1.5 block uppercase tracking-wide">{label}</label>{children}</div>
}

function ColorPicker({ value, onChange }: ColorPickerProps) {
  return (
    <div>
      <label className="text-xs font-medium text-[#8A8070] mb-2 block uppercase tracking-wide">Color</label>
      <div className="flex gap-2 flex-wrap">
        {ASSET_COLORS.map(c => (
          <button key={c} onClick={() => onChange(c)}
            className="w-7 h-7 rounded-full border-2 transition-all"
            style={{ background: c, borderColor: value === c ? '#0D0D0D' : 'transparent' }} />
        ))}
      </div>
    </div>
  )
}

// ─── SearchBar for list/history ───────────────────────────────────────────────
function SearchBar({ value, onChange, placeholder }: SearchBarProps) {
  return (
    <div className="relative">
      <Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-[#8A8070]" />
      <input
        className="w-full pl-9 pr-3 py-2 bg-[#fff] border border-[#C9A84C]/20 rounded-xl text-sm text-[#0D0D0D] placeholder-[#8A8070] focus:outline-none focus:border-[#C9A84C] transition-all"
        placeholder={placeholder}
        value={value}
        onChange={e => onChange(e.target.value)}
      />
      {value && <button onClick={() => onChange('')} className="absolute right-3 top-1/2 -translate-y-1/2 text-[#8A8070]"><X size={12} /></button>}
    </div>
  )
}

function Logo({ size = 24 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 100 100" fill="none">
      <rect width="100" height="100" rx="22" fill="bg-emerald" />
      <polyline points="14,73 32,48 50,61 67,32 85,19"
        stroke="white" strokeWidth="7" fill="none" strokeLinecap="round" strokeLinejoin="round" />
      <circle cx="67" cy="32" r="6" fill="white" />
      <circle cx="85" cy="19" r="6" fill="#FCD34D" />
    </svg>
  )
}

// ─── Main App ─────────────────────────────────────────────────────────────────
export default function App() {
  const [data, setData] = useState<AppData>(emptyData())
  const [loaded, setLoaded] = useState(false)
  const [isClient, setIsClient] = useState(false)
  const [mainTab, setMainTab] = useState('dashboard')
  const [subTab, setSubTab] = useState('list')
  const [listSearch, setListSearch] = useState('')
  const [histSearch, setHistSearch] = useState('')

  // Modals
  const [showFundModal, setShowFundModal] = useState(false)
  const [editFund, setEditFund] = useState<MutualFund | undefined>()
  const [showETFModal, setShowETFModal] = useState(false)
  const [editETF, setEditETF] = useState<ETF | undefined>()
  const [showStockModal, setShowStockModal] = useState(false)
  const [editStock, setEditStock] = useState<Stock | undefined>()
  const [showInvestModal, setShowInvestModal] = useState(false)
  const [investAssetType, setInvestAssetType] = useState('mutual-funds')
  const [editPayment, setEditPayment] = useState<Payment | undefined>()
  const [preSelectedAssetId, setPreSelectedAssetId] = useState<string | null>(null)
  const [showProjection, setShowProjection] = useState(false)

  // Per-asset history modal
  const [historyAsset, setHistoryAsset] = useState<MutualFund | ETF | Stock | null>(null)

  const fileRef = useRef<HTMLInputElement>(null)

  useEffect(() => setIsClient(true), [])
  useEffect(() => {
    if (typeof window === 'undefined') return
    const saved = localStorage.getItem(STORAGE_KEY)
    if (saved) {
      try {
        const parsed = JSON.parse(saved) as AppData
        setData(parsed)
        setLoaded(true)
      } catch (e) {
        console.error(e)
      }
    }
  }, [])

  const updateData = useCallback((fn: AppData | ((prev: AppData) => AppData)) => {
    setData(prev => {
      const next = typeof fn === 'function' ? fn(prev) : fn
      if (typeof window !== 'undefined') localStorage.setItem(STORAGE_KEY, JSON.stringify(next))
      return next
    })
  }, [])

  const handleUpload = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (!file) return
    const reader = new FileReader()
    reader.onload = ev => {
      try {
        updateData(parseExcel(ev.target?.result as ArrayBuffer))
        setLoaded(true)
      } catch {
        alert('Invalid file')
      }
    }
    reader.readAsArrayBuffer(file)
    e.target.value = ''
  }, [updateData])

  const startFresh = () => { updateData(emptyData()); setLoaded(true) }

  // CRUD
  const saveFund = (f: MutualFund) => {
    updateData(d => ({
      ...d,
      funds: d.funds.find(x => x.id === f.id) ? d.funds.map(x => x.id === f.id ? f : x) : [...d.funds, f]
    }))
    setShowFundModal(false)
    setEditFund(undefined)
  }
  const deleteFund = (id: string) => {
    if (!confirm('Delete fund and all investments?')) return
    updateData(d => ({
      ...d,
      funds: d.funds.filter(f => f.id !== id),
      payments: d.payments.filter(p => !(p.assetId === id && p.assetType === 'mutual-funds'))
    }))
  }
  const saveETF = (e: ETF) => {
    updateData(d => ({
      ...d,
      etfs: d.etfs.find(x => x.id === e.id) ? d.etfs.map(x => x.id === e.id ? e : x) : [...d.etfs, e]
    }))
    setShowETFModal(false)
    setEditETF(undefined)
  }
  const deleteETF = (id: string) => {
    if (!confirm('Delete ETF and all investments?')) return
    updateData(d => ({
      ...d,
      etfs: d.etfs.filter(e => e.id !== id),
      payments: d.payments.filter(p => !(p.assetId === id && p.assetType === 'etfs'))
    }))
  }
  const saveStock = (s: Stock) => {
    updateData(d => ({
      ...d,
      stocks: d.stocks.find(x => x.id === s.id) ? d.stocks.map(x => x.id === s.id ? s : x) : [...d.stocks, s]
    }))
    setShowStockModal(false)
    setEditStock(undefined)
  }
  const deleteStock = (id: string) => {
    if (!confirm('Delete stock and all investments?')) return
    updateData(d => ({
      ...d,
      stocks: d.stocks.filter(s => s.id !== id),
      payments: d.payments.filter(p => !(p.assetId === id && p.assetType === 'stocks'))
    }))
  }
  const savePayment = (p: Payment) => {
    updateData(d => ({
      ...d,
      payments: d.payments.find(x => x.id === p.id) ? d.payments.map(x => x.id === p.id ? p : x) : [...d.payments, p]
    }))
    setShowInvestModal(false)
    setEditPayment(undefined)
    setPreSelectedAssetId(null)
  }
  const deletePayment = (id: string) => updateData(d => ({ ...d, payments: d.payments.filter(p => p.id !== id) }))

  const openInvest = (assetType: string, assetId: string | null = null) => {
    setInvestAssetType(assetType)
    setPreSelectedAssetId(assetId)
    setEditPayment(undefined)
    setShowInvestModal(true)
  }

  // Open asset history from the history modal (edit)
  const openEditFromHistory = (payment: Payment) => {
    setHistoryAsset(null)
    setEditPayment(payment)
    setInvestAssetType(payment.assetType)
    setPreSelectedAssetId(null)
    setShowInvestModal(true)
  }

  // Derived
  const totalInvested = calcTotalInvested(data.payments)
  const totalMonthly = [...data.funds, ...data.etfs].reduce((s, a) => s + a.sipAmount, 0)
  const totalStockValue = data.stocks.reduce((s, s2) => s + (s2.quantity * s2.buyPrice), 0)

  const getCurrentAssets = (): (MutualFund | ETF | Stock)[] => {
    if (mainTab === 'mutual-funds') return data.funds
    if (mainTab === 'etfs') return data.etfs
    if (mainTab === 'stocks') return data.stocks
    return []
  }
  const getCurrentPayments = (): Payment[] => data.payments.filter(p => p.assetType === mainTab)

  // Filtered assets (list tab search)
  const filteredAssets = useMemo(() => {
    const assets = getCurrentAssets()
    if (!listSearch) return assets
    const lq = listSearch.toLowerCase()
    return assets.filter(a =>
      a.name.toLowerCase().includes(lq) ||
      a.category?.toLowerCase().includes(lq) ||
      ('symbol' in a && a.symbol?.toLowerCase().includes(lq))
    )
  }, [mainTab, data, listSearch])

  // Filtered payments (history tab search)
  const filteredPayments = useMemo(() => {
    const payments = getCurrentPayments()
    if (!histSearch) return payments
    const lq = histSearch.toLowerCase()
    const assets = getCurrentAssets()
    return payments.filter(p => {
      const asset = assets.find(a => a.id === p.assetId)
      return asset?.name.toLowerCase().includes(lq) ||
        p.date.includes(lq) ||
        getMonthLabel(p.date).toLowerCase().includes(lq) ||
        p.notes?.toLowerCase().includes(lq)
    })
  }, [mainTab, data, histSearch])

  const pieData = [...data.funds, ...data.etfs, ...data.stocks].map(a => ({
    name: a.name,
    value: calcTotalInvested(data.payments, a.id, a.assetType),
    color: a.color
  })).filter(x => x.value > 0)

  const allMonths = [...new Set(data.payments.map(p => p.date))].sort()
  const timelineData = allMonths.map(m => {
    const mp = data.payments.filter(p => p.date === m)
    return {
      month: getMonthLabel(m),
      mutualFunds: mp.filter(p => p.assetType === 'mutual-funds').reduce((s, p) => s + p.amount, 0),
      etfs: mp.filter(p => p.assetType === 'etfs').reduce((s, p) => s + p.amount, 0),
      stocks: mp.filter(p => p.assetType === 'stocks').reduce((s, p) => s + p.amount, 0)
    }
  })

  if (!isClient) {
    return (
      <div className="min-h-screen bg-paper flex flex-col items-center justify-center p-6">
        <div className="max-w-lg w-full text-center">
          <div className="w-16 h-16 rounded-2xl bg-emerald flex items-center justify-center mx-auto mb-6 animate-pulse">
            <Logo size={44} />
          </div>
        </div>
      </div>
    )
  }
  // ── Landing ──
  if (!loaded) {
    const hasSaved = typeof window !== 'undefined' ? localStorage.getItem(STORAGE_KEY) : null
    return (
      <div className="min-h-screen bg-[#fff] flex flex-col items-center justify-center p-6">
        <div className="max-w-sm w-full text-center">
          <div className="w-14 h-14 rounded-2xl bg-[#1A5C3A] flex items-center justify-center mx-auto mb-5">
            <TrendingUp size={26} className="text-[#FDF8F0]" />
          </div>
          <h1 className="text-3xl font-bold text-[#0D0D0D] mb-2">Investment Tracker</h1>
          <p className="text-[#8A8070] text-sm mb-8">Track Mutual Funds, ETFs & Stocks. Data saved in your browser.</p>
          {hasSaved && <button className="btn-primary w-full mb-4" onClick={() => setLoaded(true)}><RefreshCw size={14} className="inline mr-2" />Continue with Saved Data</button>}
          <div className="bg-[#FDF8F0] rounded-2xl border border-[#C9A84C]/20 p-5 mb-4">
            <p className="text-sm font-medium text-[#0D0D0D] mb-3">Load from Excel</p>
            <input ref={fileRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={handleUpload} />
            <button className="btn-gold w-full" onClick={() => fileRef.current?.click()}><Upload size={14} className="inline mr-2" />Upload Excel File</button>
          </div>
          <div className="relative my-3"><div className="absolute inset-0 flex items-center"><div className="w-full border-t border-[#C9A84C]/20" /></div><span className="relative bg-[#fff] px-3 text-xs text-[#8A8070]">or</span></div>
          <button className="btn-outline w-full" onClick={startFresh}><Plus size={14} className="inline mr-2" />Start Fresh</button>
        </div>
      </div>
    )
  }

  const currentAssets = getCurrentAssets()
  const currentPayments = getCurrentPayments()
  const isMFTab = mainTab === 'mutual-funds'
  const isETFTab = mainTab === 'etfs'
  const isStockTab = mainTab === 'stocks'

  return (
    <div className="min-h-screen bg-[#fff]">

      {/* ── Header ── */}
      <header className="bg-[#fff] border-b border-[#C9A84C]/15 sticky top-0 z-30 shadow-sm">
        <div className="max-w-6xl mx-auto px-3 sm:px-4 py-3 flex items-center justify-between gap-2">
          <div className="flex items-center gap-2 shrink-0">
            <div className="w-8 h-8 rounded-xl bg-[#1A5C3A] flex items-center justify-center">
              <Logo />
            </div>
            <span className="font-bold text-[#0D0D0D] text-sm sm:block">TrackMyFund</span>
          </div>

          <div className="flex items-center gap-1.5 shrink-0">
            <input ref={fileRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={handleUpload} />
            <button className="btn-outline text-xs py-2 px-2.5" onClick={() => fileRef.current?.click()}><Upload size={13} /><span className="hidden sm:inline">Load</span></button>
            <button className="btn-gold text-xs py-2 px-2.5" onClick={() => exportExcel(data)}><Download size={13} /><span className="sm:inline">Export</span></button>
          </div>
        </div>

        {/* Main Tabs */}
        <div className="max-w-6xl mx-auto px-3 sm:px-4 flex overflow-x-auto overflow-y-hidden">
          {([
            ['dashboard', 'Dashboard', BarChart3],
            ['mutual-funds', 'Mutual Funds', Wallet],
            ['etfs', 'ETFs', BarChart2],
            ['stocks', 'Stocks', Activity]
          ] as const).map(([id, label, Icon]) => (
            <button
              key={id}
              onClick={() => {
                setMainTab(id);
                setSubTab('list');
                setListSearch('');
                setHistSearch('');
              }}
              className={`flex items-center gap-1.5 px-3 sm:px-4 py-2.5 text-xs sm:text-sm font-medium border-b-2 transition-all shrink-0 -mb-px whitespace-nowrap ${mainTab === id
                ? 'border-[#1A5C3A] text-[#1A5C3A]'
                : 'border-transparent text-[#8A8070] hover:text-[#0D0D0D]'
                }`}
            >
              <Icon size={13} />
              {label}
            </button>
          ))}
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-3 sm:px-4 py-4 sm:py-3">

        {/* ── DASHBOARD ── */}
        {mainTab === 'dashboard' && (
          <div className="space-y-5">
            <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-5 gap-2 sm:gap-3">
              {[
                { l: 'Total Invested', v: `₹${totalInvested.toLocaleString('en-IN')}`, c: 'text-[#1A5C3A]' },
                { l: 'Monthly SIP', v: `₹${totalMonthly.toLocaleString('en-IN')}`, c: 'text-[#8B6914]' },
                { l: 'Stock Value', v: `₹${totalStockValue.toLocaleString('en-IN')}`, c: 'text-[#1565C0]' },
                { l: 'Active Assets', v: (data.funds.length + data.etfs.length + data.stocks.length).toString(), c: 'text-[#0D0D0D]' },
                { l: 'Months Tracked', v: [...new Set(data.payments.map(p => p.date))].length.toString(), c: 'text-[#0D0D0D]' },
              ].map(s => (
                <div key={s.l} className="card p-3 sm:p-4 flex flex-col gap-0.5">
                  <span className="text-xs text-[#8A8070]">{s.l}</span>
                  <span className={`font-bold text-lg sm:text-xl ${s.c}`}>{s.v}</span>
                </div>
              ))}
            </div>

            {data.funds.length === 0 && data.etfs.length === 0 && data.stocks.length === 0 ? (
              <div className="card p-10 text-center">
                <TrendingUp size={36} className="text-[#C9A84C]/40 mx-auto mb-3" />
                <p className="font-medium text-[#0D0D0D] mb-1">No investments yet</p>
                <p className="text-sm text-[#8A8070] mb-5">Add mutual funds, ETFs or stocks to get started</p>
                <div className="flex gap-2 justify-center flex-wrap">
                  {(['mutual-funds', 'etfs', 'stocks'] as const).map(t => (
                    <button key={t} className="btn-outline text-xs" onClick={() => setMainTab(t)}>
                      Add {t === 'mutual-funds' ? 'Fund' : t === 'etfs' ? 'ETF' : 'Stock'}
                    </button>
                  ))}
                </div>
              </div>
            ) : (
              <div className="grid md:grid-cols-2 gap-4">
                <div className="card p-4 sm:p-5">
                  <h3 className="font-semibold text-[#0D0D0D] mb-4">Portfolio Allocation</h3>
                  {pieData.length === 0 ? <p className="text-sm text-[#8A8070] text-center py-8">Log investments to see allocation</p> : (
                    <ResponsiveContainer width="100%" height={200}>
                      <PieChart>
                        <Pie data={pieData} cx="50%" cy="50%" outerRadius={75} dataKey="value"
                          label={({ name, percent }) => `${name.split(' ').pop()} ${(percent * 100).toFixed(0)}%`} labelLine={false}>
                          {pieData.map((e, i) => <Cell key={i} fill={e.color} />)}
                        </Pie>
                        <Tooltip formatter={(v: number) => [`₹${Number(v).toLocaleString('en-IN')}`, 'Invested']} contentStyle={{ background: '#FDF8F0', border: '1px solid #C9A84C33', borderRadius: 12 }} />
                      </PieChart>
                    </ResponsiveContainer>
                  )}
                </div>
                <div className="card p-4 sm:p-5">
                  <h3 className="font-semibold text-[#0D0D0D] mb-4">Monthly by Type</h3>
                  {timelineData.length === 0 ? <p className="text-sm text-[#8A8070] text-center py-8">Log investments to see chart</p> : (
                    <ResponsiveContainer width="100%" height={200}>
                      <BarChart data={timelineData.slice(-12)}>
                        <XAxis dataKey="month" tick={{ fontSize: 10, fill: '#8A8070' }} />
                        <YAxis tickFormatter={v => `₹${(v / 1000).toFixed(0)}k`} tick={{ fontSize: 10, fill: '#8A8070' }} />
                        <Tooltip formatter={(v: number) => [`₹${Number(v).toLocaleString('en-IN')}`, '']} contentStyle={{ background: '#FDF8F0', border: '1px solid #C9A84C33', borderRadius: 12 }} />
                        <Bar dataKey="mutualFunds" stackId="a" fill="#C9A84C" name="Mutual Funds" />
                        <Bar dataKey="etfs" stackId="a" fill="#1A5C3A" name="ETFs" />
                        <Bar dataKey="stocks" stackId="a" fill="#1565C0" name="Stocks" />
                      </BarChart>
                    </ResponsiveContainer>
                  )}
                </div>
              </div>
            )}

            {[...data.funds, ...data.etfs, ...data.stocks].length > 0 && (
              <div>
                <div className="flex items-center justify-between mb-3">
                  <h3 className="font-semibold text-[#0D0D0D]">Asset Overview</h3>
                  <button className="btn-outline text-xs py-2 px-3" onClick={() => setShowProjection(true)}><TrendingUp size={13} />Projection</button>
                </div>
                <div className="grid sm:grid-cols-2 gap-3">
                  {[...data.funds, ...data.etfs, ...data.stocks].slice(0, 6).map(asset => {
                    const at = asset.assetType
                    const invested = calcTotalInvested(data.payments, asset.id, at)
                    const monthlyAmt = isMutualFund(asset) || isETF(asset) ? asset.sipAmount : 0
                    const p10y = calcFutureValue(monthlyAmt, 120, asset.expectedReturn)
                    return (
                      <div key={asset.id} className="card p-4 flex gap-3">
                        <div className="w-1 rounded-full shrink-0 self-stretch" style={{ background: asset.color }} />
                        <div className="flex-1 min-w-0">
                          <div className="flex justify-between gap-2">
                            <div className="min-w-0">
                              <p className="font-medium text-sm text-[#0D0D0D] truncate">{asset.name}</p>
                              <p className="text-xs text-[#8A8070]">{asset.category} · {asset.expectedReturn}% CAGR</p>
                            </div>
                            <div className="text-right shrink-0">
                              <p className="text-xs text-[#8A8070]">Invested</p>
                              <p className="font-semibold text-[#1A5C3A] text-sm">₹{invested.toLocaleString('en-IN')}</p>
                            </div>
                          </div>
                          {(isMutualFund(asset) || isETF(asset)) && (
                            <div className="flex gap-3 mt-2 text-xs text-[#8A8070]">
                              <span>SIP: <strong className="text-[#0D0D0D]">₹{asset.sipAmount.toLocaleString('en-IN')}/mo</strong></span>
                              <span>10Y: <strong className="text-[#8B6914]">₹{(p10y / 100000).toFixed(1)}L</strong></span>
                            </div>
                          )}
                          {isStock(asset) && (
                            <div className="flex gap-3 mt-2 text-xs text-[#8A8070]">
                              <span>Qty: <strong className="text-[#0D0D0D]">{asset.quantity}</strong></span>
                              <span>Avg: <strong className="text-[#0D0D0D]">₹{asset.buyPrice.toLocaleString('en-IN')}</strong></span>
                            </div>
                          )}
                        </div>
                      </div>
                    )
                  })}
                </div>
              </div>
            )}
          </div>
        )}

        {/* ── ASSET MANAGEMENT ── */}
        {mainTab !== 'dashboard' && (
          <div className="space-y-4">
            {/* Sub-tab bar + Add button */}
            <div className="flex items-center justify-between gap-3">
              <div className="flex border-b border-[#C9A84C]/20">
                {[
                  ['list', 'Assets'],
                  ['investments', 'History']
                ].map(([id, label]) => (
                  <button key={id} onClick={() => { setSubTab(id); setListSearch(''); setHistSearch('') }}
                    className={`px-3 sm:px-4 py-2 text-sm font-medium transition-all whitespace-nowrap ${subTab === id ? 'text-[#1A5C3A] border-b-2 border-[#1A5C3A]' : 'text-[#8A8070] hover:text-[#0D0D0D]'}`}>
                    {label}
                    {id === 'investments' && currentPayments.length > 0 && <span className="ml-1.5 text-xs bg-[#1A5C3A]/10 text-[#1A5C3A] px-1.5 py-0.5 rounded-full">{currentPayments.length}</span>}
                  </button>
                ))}
              </div>
              <div className="flex gap-2">
                {subTab === 'investments' && currentAssets.length > 0 && (
                  <button className="btn-primary text-xs py-2 px-2" onClick={() => openInvest(mainTab)}>
                    <Plus size={13} />Log
                  </button>
                )}
                <button className="btn-primary text-xs py-2 px-2" onClick={() => {
                  if (isMFTab) setShowFundModal(true)
                  else if (isETFTab) setShowETFModal(true)
                  else setShowStockModal(true)
                }}>
                  <Plus size={14} />Add {isMFTab ? 'Fund' : isETFTab ? 'ETF' : 'Stock'}
                </button>
              </div>
            </div>

            {/* ── ASSETS LIST ── */}
            {subTab === 'list' && (
              <div className="space-y-3">
                {currentAssets.length > 0 && (
                  <SearchBar value={listSearch} onChange={setListSearch} placeholder={`Search ${isMFTab ? 'funds' : isETFTab ? 'ETFs' : 'stocks'} by name, category...`} />
                )}
                {currentAssets.length === 0 ? (
                  <div className="card p-10 text-center">
                    <Wallet size={36} className="text-[#C9A84C]/40 mx-auto mb-3" />
                    <p className="font-medium text-[#0D0D0D] mb-1">No {isMFTab ? 'funds' : isETFTab ? 'ETFs' : 'stocks'} added</p>
                    <p className="text-sm text-[#8A8070] mb-4">Add your first to start tracking</p>
                    <button className="btn-primary mx-auto" onClick={() => {
                      if (isMFTab) setShowFundModal(true)
                      else if (isETFTab) setShowETFModal(true)
                      else setShowStockModal(true)
                    }}>
                      <Plus size={14} />Add First {isMFTab ? 'Fund' : isETFTab ? 'ETF' : 'Stock'}
                    </button>
                  </div>
                ) : filteredAssets.length === 0 ? (
                  <div className="card p-8 text-center">
                    <Search size={28} className="text-[#C9A84C]/40 mx-auto mb-2" />
                    <p className="text-sm text-[#8A8070]">No results for "{listSearch}"</p>
                  </div>
                ) : (
                  filteredAssets.map(asset => {
                    const invested = calcTotalInvested(data.payments, asset.id, mainTab)
                    const investments = data.payments.filter(p => p.assetId === asset.id && p.assetType === mainTab)
                    const totalUnits = investments.reduce((s, p) => s + (p.units || 0), 0)
                    const totalQty = investments.reduce((s, p) => s + (p.quantity || 0), 0)
                    // Avg NAV = total invested / total units (weighted average cost)
                    const avgNav = isMFTab && totalUnits > 0 ? invested / totalUnits : null
                    const lastPrice = [...investments].sort((a, b) => b.date.localeCompare(a.date))[0]?.price

                    return (
                      <div key={asset.id} className="card p-4 sm:p-5">
                        <div className="flex items-start gap-3">
                          <div className="w-9 h-9 rounded-xl flex items-center justify-center shrink-0" style={{ background: asset.color + '22' }}>
                            <TrendingUp size={16} style={{ color: asset.color }} />
                          </div>
                          <div className="flex-1 min-w-0">
                            <div className="flex items-start justify-between gap-2">
                              <div className="min-w-0">
                                <h3 className="font-semibold text-[#0D0D0D] text-sm sm:text-base truncate">{asset.name}</h3>
                                {'symbol' in asset && <p className="text-xs text-[#8A8070]">{asset.symbol}</p>}
                                <p className="text-xs text-[#8A8070] mt-0.5">{asset.category} · Since {asset.startDate} · {asset.expectedReturn}% CAGR</p>
                              </div>
                              <div className="flex gap-1 shrink-0">
                                <button
                                  onClick={() => openInvest(mainTab, asset.id)}
                                  className="p-1.5 text-[#1A5C3A] hover:bg-[#1A5C3A]/10 rounded-lg transition-all"
                                  title="Log Investment"
                                >
                                  <Plus size={14} />
                                </button>
                                <button
                                  onClick={() => setHistoryAsset(asset)}
                                  className="p-1.5 text-[#8A8070] hover:text-[#1565C0] hover:bg-[#1565C0]/10 rounded-lg transition-all"
                                  title="View History"
                                >
                                  <Clock size={14} />
                                </button>
                                <button onClick={() => {
                                  if (isMFTab) { setEditFund(asset as MutualFund); setShowFundModal(true) }
                                  else if (isETFTab) { setEditETF(asset as ETF); setShowETFModal(true) }
                                  else { setEditStock(asset as Stock); setShowStockModal(true) }
                                }} className="p-1.5 text-[#8A8070] hover:text-[#0D0D0D] rounded-lg hover:bg-[#F5F0E8] transition-all">
                                  <Edit3 size={14} />
                                </button>
                                <button onClick={() => {
                                  if (isMFTab) deleteFund(asset.id)
                                  else if (isETFTab) deleteETF(asset.id)
                                  else deleteStock(asset.id)
                                }} className="p-1.5 text-[#8A8070] hover:text-red-500 rounded-lg hover:bg-[#F5F0E8] transition-all">
                                  <Trash2 size={14} />
                                </button>
                              </div>
                            </div>

                            <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 mt-3">
                              {(isMutualFund(asset) || isETF(asset)) && (
                                <div className="bg-[#fff] rounded-xl p-2.5 border shadow-sm">
                                  <p className="text-xs text-[#8A8070]">Monthly SIP</p>
                                  <p className="font-bold text-[#1A5C3A] text-sm">₹{asset.sipAmount.toLocaleString('en-IN')}</p>
                                </div>
                              )}
                              <div className="bg-[#fff] rounded-xl p-2.5 border shadow-sm">
                                <p className="text-xs text-[#8A8070]">Total Invested</p>
                                <p className="font-bold text-[#0D0D0D] text-sm">₹{invested.toLocaleString('en-IN')}</p>
                              </div>
                              <div className="bg-[#fff] rounded-xl p-2.5 border shadow-sm">
                                <p className="text-xs text-[#8A8070]">Entries</p>
                                <p className="font-bold text-[#0D0D0D] text-sm">{investments.length}</p>
                              </div>

                              {isMFTab && avgNav !== null && (
                                <div className="bg-[#fff] rounded-xl p-2.5 border shadow-sm">
                                  <p className="text-xs text-[#8A8070]">Avg NAV</p>
                                  <p className="font-bold text-[#8B6914] text-sm">₹{avgNav.toFixed(3)}</p>
                                </div>
                              )}
                              {isMFTab && totalUnits > 0 && (
                                <div className="bg-[#fff] rounded-xl p-2.5 border shadow-sm">
                                  <p className="text-xs text-[#8A8070]">Total Units</p>
                                  <p className="font-bold text-[#4A148C] text-sm">{totalUnits.toFixed(3)}</p>
                                </div>
                              )}

                              {isETFTab && lastPrice && (
                                <div className="bg-[#fff] rounded-xl p-2.5 border shadow-sm">
                                  <p className="text-xs text-[#8A8070]">Last Price</p>
                                  <p className="font-bold text-[#8B6914] text-sm">₹{lastPrice.toFixed(2)}</p>
                                </div>
                              )}
                              {isETFTab && totalQty > 0 && (
                                <div className="bg-[#fff] rounded-xl p-2.5 border shadow-sm">
                                  <p className="text-xs text-[#8A8070]">Total Units</p>
                                  <p className="font-bold text-[#4A148C] text-sm">{totalQty}</p>
                                </div>
                              )}

                              {isStock(asset) && (
                                <>
                                  <div className="bg-[#fff] rounded-xl p-2.5 border shadow-sm">
                                    <p className="text-xs text-[#8A8070]">Quantity</p>
                                    <p className="font-bold text-[#0D0D0D] text-sm">{totalQty > 0 ? totalQty : asset.quantity}</p>
                                  </div>
                                  <div className="bg-[#fff] rounded-xl p-2.5 border shadow-sm">
                                    <p className="text-xs text-[#8A8070]">Avg Buy Price</p>
                                    <p className="font-bold text-[#8B6914] text-sm">₹{asset.buyPrice.toLocaleString('en-IN')}</p>
                                  </div>
                                  {lastPrice && (
                                    <div className="bg-[#fff] rounded-xl p-2.5 border shadow-sm">
                                      <p className="text-xs text-[#8A8070]">Last Price</p>
                                      <p className="font-bold text-[#1565C0] text-sm">₹{lastPrice.toLocaleString('en-IN')}</p>
                                    </div>
                                  )}
                                </>
                              )}
                            </div>
                          </div>
                        </div>
                      </div>
                    )
                  })
                )}
              </div>
            )}

            {/* ── INVESTMENT HISTORY ── */}
            {subTab === 'investments' && (
              <div className="space-y-3">
                {currentAssets.length === 0 ? (
                  <div className="card p-10 text-center">
                    <Calendar size={36} className="text-[#C9A84C]/40 mx-auto mb-3" />
                    <p className="font-medium text-[#0D0D0D] mb-1">No assets added yet</p>
                    <p className="text-sm text-[#8A8070]">Add assets before logging investments</p>
                  </div>
                ) : currentPayments.length === 0 ? (
                  <div className="card p-10 text-center">
                    <Calendar size={36} className="text-[#C9A84C]/40 mx-auto mb-3" />
                    <p className="font-medium text-[#0D0D0D] mb-1">No investments logged yet</p>
                    <button className="btn-primary mx-auto mt-4" onClick={() => openInvest(mainTab)}><Plus size={14} />Log First Investment</button>
                  </div>
                ) : (
                  <>
                    <SearchBar value={histSearch} onChange={setHistSearch} placeholder="Search by fund name, month (e.g. Jun), notes..." />
                    {filteredPayments.length === 0 ? (
                      <div className="card p-8 text-center">
                        <Search size={28} className="text-[#C9A84C]/40 mx-auto mb-2" />
                        <p className="text-sm text-[#8A8070]">No results for "{histSearch}"</p>
                      </div>
                    ) : (
                      <>
                        {/* Mobile view */}
                        <div className="sm:hidden space-y-2">
                          {[...filteredPayments].sort((a, b) => b.date.localeCompare(a.date)).map(p => {
                            const asset = currentAssets.find(a => a.id === p.assetId)
                            return (
                              <div key={p.id} className="card p-3">
                                <div className="flex items-start justify-between gap-2">
                                  <div className="flex items-center gap-2 min-w-0">
                                    <div className="w-2 h-2 rounded-full shrink-0 mt-1" style={{ background: asset?.color ?? '#ccc' }} />
                                    <div className="min-w-0">
                                      <p className="text-sm font-medium text-[#0D0D0D] truncate">{asset?.name ?? 'Unknown'}</p>
                                      <p className="text-xs text-[#8A8070]">{getMonthLabel(p.date)}</p>
                                    </div>
                                  </div>
                                  <div className="text-right shrink-0">
                                    <p className="font-semibold text-[#1A5C3A] text-sm">₹{p.amount.toLocaleString('en-IN')}</p>
                                  </div>
                                </div>
                                {(p.nav || p.units || p.quantity || p.price) && (
                                  <div className="flex flex-wrap gap-2 mt-2">
                                    {p.nav && <span className="text-xs bg-[#fff] text-[#8A8070] px-2 py-0.5 rounded-lg">NAV ₹{p.nav}</span>}
                                    {p.units && <span className="text-xs bg-[#fff] text-[#4A148C] px-2 py-0.5 rounded-lg">{p.units} units</span>}
                                    {p.price && !p.nav && <span className="text-xs bg-[#fff] text-[#8A8070] px-2 py-0.5 rounded-lg">₹{p.price}/unit</span>}
                                    {p.quantity && <span className="text-xs bg-[#fff] text-[#0D0D0D] px-2 py-0.5 rounded-lg">Qty {p.quantity}</span>}
                                  </div>
                                )}
                                <div className="flex items-center justify-between mt-2">
                                  <span className="text-xs text-[#8A8070] truncate">{p.notes || ''}</span>
                                  <div className="flex gap-1">
                                    <button onClick={() => { setEditPayment(p); setInvestAssetType(mainTab); setPreSelectedAssetId(null); setShowInvestModal(true) }} className="p-1.5 text-[#8A8070] hover:text-[#0D0D0D] rounded-lg"><Edit3 size={12} /></button>
                                    <button onClick={() => deletePayment(p.id)} className="p-1.5 text-[#8A8070] hover:text-red-500 rounded-lg"><Trash2 size={12} /></button>
                                  </div>
                                </div>
                              </div>
                            )
                          })}
                          <div className="card p-3 flex justify-between">
                            <span className="text-xs font-medium text-[#8A8070]">Total ({filteredPayments.length} entries)</span>
                            <span className="font-bold text-[#1A5C3A]">₹{filteredPayments.reduce((s, p) => s + p.amount, 0).toLocaleString('en-IN')}</span>
                          </div>
                        </div>

                        {/* Desktop table */}
                        <div className="hidden sm:block card overflow-hidden">
                          <div className="overflow-x-auto scrollbar-thin">
                            <table className="w-full text-sm">
                              <thead>
                                <tr className="border-b border-[#C9A84C]/15 bg-[#F5F0E8]/50">
                                  <th className="text-left px-4 py-3 text-sm font-semibold text-[#4a4a4a] tracking-wide">Month</th>
                                  <th className="text-left px-4 py-3 text-sm font-semibold text-[#4a4a4a] tracking-wide">Name</th>
                                  <th className="text-right px-4 py-3 text-sm font-semibold text-[#4a4a4a] tracking-wide">Amount</th>
                                  {isMFTab && <><th className="text-right px-4 py-3 text-sm font-semibold text-[#4a4a4a]  tracking-wide">NAV</th><th className="text-right px-4 py-3 text-sm font-semibold text-[#4a4a4a]  tracking-wide">Units</th></>}
                                  {(isETFTab || isStockTab) && <><th className="text-right px-4 py-3 text-sm font-semibold text-[#4a4a4a]  tracking-wide">Price</th><th className="text-right px-4 py-3 text-sm font-semibold text-[#4a4a4a]  tracking-wide">Qty</th></>}
                                  <th className="text-left px-4 py-3 text-sm font-semibold text-[#4a4a4a] tracking-wide">Notes</th>
                                  <th className="px-4 py-3" />
                                </tr>
                              </thead>
                              <tbody>
                                {[...filteredPayments].sort((a, b) => b.date.localeCompare(a.date)).map(p => {
                                  const asset = currentAssets.find(a => a.id === p.assetId)
                                  return (
                                    <tr key={p.id} className="border-b border-[#C9A84C]/10 hover:bg-[#F5F0E8]/40 transition-colors">
                                      <td className="px-4 py-3 text-sm text-[#4a4a4a]">{getMonthLabel(p.date)}</td>
                                      <td className="px-4 py-3">
                                        <div className="flex items-center gap-2">
                                          <div className="w-2 h-2 rounded-full shrink-0" style={{ background: asset?.color ?? '#ccc' }} />
                                          <span className="truncate max-w-[160px]">{asset?.name ?? 'Unknown'}</span>
                                        </div>
                                      </td>
                                      <td className="px-4 py-3 text-right font-medium text-[#1A5C3A]">₹{p.amount.toLocaleString('en-IN')}</td>
                                      {isMFTab && <>
                                        <td className="px-4 py-3 text-right text-[#4a4a4a]">{p.nav ? `₹${p.nav}` : '—'}</td>
                                        <td className="px-4 py-3 text-right text-[#4A148C] font-medium">{p.units || '—'}</td>
                                      </>}
                                      {(isETFTab || isStockTab) && <>
                                        <td className="px-4 py-3 text-right text-[#4a4a4a]">{p.price ? `₹${p.price}` : '—'}</td>
                                        <td className="px-4 py-3 text-right text-[#4a4a4a]">{p.quantity || '—'}</td>
                                      </>}
                                      <td className="px-4 py-3 text-[#4a4a4a] text-xs truncate max-w-[100px]">{p.notes || '—'}</td>
                                      <td className="px-4 py-3">
                                        <div className="flex gap-1 justify-end">
                                          <button onClick={() => { setEditPayment(p); setInvestAssetType(mainTab); setPreSelectedAssetId(null); setShowInvestModal(true) }} className="p-1.5 text-[#8A8070] hover:text-[#0D0D0D] rounded-lg hover:bg-[#F5F0E8] transition-all"><Edit3 size={13} /></button>
                                          <button onClick={() => deletePayment(p.id)} className="p-1.5 text-[#8A8070] hover:text-red-500 rounded-lg hover:bg-[#F5F0E8] transition-all"><Trash2 size={13} /></button>
                                        </div>
                                      </td>
                                    </tr>
                                  )
                                })}
                              </tbody>
                              <tfoot>
                                <tr className="bg-[#fff]/50">
                                  <td colSpan={2} className="px-4 py-3 text-xs font-medium text-[#8A8070]">Total ({filteredPayments.length} entries)</td>
                                  <td className="px-4 py-3 text-right font-bold text-[#1A5C3A]">₹{filteredPayments.reduce((s, p) => s + p.amount, 0).toLocaleString('en-IN')}</td>
                                  <td colSpan={99} />
                                </tr>
                              </tfoot>
                            </table>
                          </div>
                        </div>
                      </>
                    )}
                  </>
                )}
              </div>
            )}
          </div>
        )}
      </main>

      {/* ── Modals ── */}
      {showFundModal && <FundModal fund={editFund} onSave={saveFund} onClose={() => { setShowFundModal(false); setEditFund(undefined) }} />}
      {showETFModal && <ETFModal etf={editETF} onSave={saveETF} onClose={() => { setShowETFModal(false); setEditETF(undefined) }} />}
      {showStockModal && <StockModal stock={editStock} onSave={saveStock} onClose={() => { setShowStockModal(false); setEditStock(undefined) }} />}
      {showInvestModal && (
        <InvestmentModal
          assetType={investAssetType}
          funds={data.funds}
          etfs={data.etfs}
          stocks={data.stocks}
          payment={editPayment}
          preSelectedAssetId={preSelectedAssetId}
          onSave={savePayment}
          onClose={() => { setShowInvestModal(false); setEditPayment(undefined); setPreSelectedAssetId(null) }}
        />
      )}
      {showProjection && <ProjectionModal funds={data.funds} etfs={data.etfs} stocks={data.stocks} onClose={() => setShowProjection(false)} />}
      {historyAsset && (
        <AssetHistoryModal
          asset={historyAsset}
          payments={data.payments.filter(p => p.assetId === historyAsset.id && p.assetType === historyAsset.assetType)}
          onClose={() => setHistoryAsset(null)}
          onEdit={openEditFromHistory}
          onDelete={(id) => { deletePayment(id) }}
        />
      )}

      <p className="text-center text-[#8A8070] text-xs pb-4">Copyright © 2026 Sudeep Teja.</p>
    </div>
  )
}