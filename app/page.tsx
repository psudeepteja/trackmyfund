'use client'

import { useState, useCallback, useRef, useEffect } from 'react'
import {
  Upload, Download, Plus, Trash2, TrendingUp, Wallet,
  Calendar, BarChart3, ChevronDown, ChevronUp, RefreshCw,
  FileSpreadsheet, Info, X, Edit3, Check, BarChart2, Activity
} from 'lucide-react'
import {
  AreaChart, Area, XAxis, YAxis, Tooltip, ResponsiveContainer,
  PieChart, Pie, Cell, Legend, BarChart, Bar
} from 'recharts'
import {
  parseExcel, exportExcel, emptyData, generateId, currentYearMonth,
  calcFutureValue, calcTotalInvested, getMonthLabel,
  Fund, ETF, Stock, Payment, TrackerData, AssetType
} from '@/lib/excel'

const CATEGORIES = ['Equity', 'Debt', 'Hybrid', 'ELSS', 'Index', 'Gold', 'International', 'Other']
const ETF_CATEGORIES = ['Large Cap', 'Mid Cap', 'Small Cap', 'Sectoral', 'Gold', 'International', 'Debt', 'Other']
const STOCK_CATEGORIES = ['Large Cap', 'Mid Cap', 'Small Cap', 'Dividend', 'Growth', 'PSU', 'Banking', 'IT', 'Pharma', 'Other']
const ASSET_COLORS = ['#C9A84C', '#1A5C3A', '#8B4513', '#4A148C', '#1565C0', '#BF360C', '#00695C', '#6A1B9A', '#E91E63', '#9C27B0']

// Storage key for localStorage
const STORAGE_KEY = 'investment-tracker-data'

// ─── Modal ───────────────────────────────────────────────────────────────────

function Modal({ title, onClose, children }: { title: string; onClose: () => void; children: React.ReactNode }) {
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4" onClick={onClose}>
      <div className="absolute inset-0 bg-ink/40 backdrop-blur-sm" />
      <div className="relative bg-cream rounded-2xl shadow-2xl border border-gold/20 w-full max-w-md p-6"
        onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between mb-5">
          <h3 className=" text-lg font-semibold text-ink">{title}</h3>
          <button onClick={onClose} className="text-muted hover:text-ink transition-colors"><X size={18} /></button>
        </div>
        {children}
      </div>
    </div>
  )
}

// ─── Add/Edit Asset Modals ──────────────────────────────────────────────────────

function FundModal({ fund, onSave, onClose }: {
  fund?: Fund; onSave: (f: Fund) => void; onClose: () => void
}) {
  const [form, setForm] = useState<Omit<Fund, 'id'>>({
    name: fund?.name ?? '',
    category: fund?.category ?? 'Equity',
    sipAmount: fund?.sipAmount ?? 1000,
    startDate: fund?.startDate ?? currentYearMonth(),
    expectedReturn: fund?.expectedReturn ?? 12,
    color: fund?.color ?? ASSET_COLORS[0],
  })

  const set = (k: keyof typeof form, v: any) => setForm((p) => ({ ...p, [k]: v }))

  const handleSave = () => {
    if (!form.name.trim()) return
    onSave({ ...form, id: fund?.id ?? generateId() })
  }

  return (
    <Modal title={fund ? 'Edit Mutual Fund' : 'Add New Mutual Fund'} onClose={onClose}>
      <div className="space-y-4">
        <div>
          <label className="text-xs font-medium text-muted mb-1.5 block uppercase tracking-wide">Fund Name</label>
          <input className="input-field" placeholder="e.g. Mirae Asset Large Cap" value={form.name}
            onChange={(e) => set('name', e.target.value)} />
        </div>
        <div className="grid grid-cols-2 gap-3">
          <div>
            <label className="text-xs font-medium text-muted mb-1.5 block uppercase tracking-wide">Category</label>
            <select className="input-field" value={form.category} onChange={(e) => set('category', e.target.value)}>
              {CATEGORIES.map((c) => <option key={c}>{c}</option>)}
            </select>
          </div>
          <div>
            <label className="text-xs font-medium text-muted mb-1.5 block uppercase tracking-wide">Monthly SIP (₹)</label>
            <input className="input-field" type="number" min={100} step={100} value={form.sipAmount}
              onChange={(e) => set('sipAmount', Number(e.target.value))} />
          </div>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <div>
            <label className="text-xs font-medium text-muted mb-1.5 block uppercase tracking-wide">Start Month</label>
            <input className="input-field" type="month" value={form.startDate}
              onChange={(e) => set('startDate', e.target.value)} />
          </div>
          <div>
            <label className="text-xs font-medium text-muted mb-1.5 block uppercase tracking-wide">Expected CAGR %</label>
            <input className="input-field" type="number" min={1} max={50} step={0.5} value={form.expectedReturn}
              onChange={(e) => set('expectedReturn', Number(e.target.value))} />
          </div>
        </div>
        <div>
          <label className="text-xs font-medium text-muted mb-2 block uppercase tracking-wide">Color</label>
          <div className="flex gap-2 flex-wrap">
            {ASSET_COLORS.map((c) => (
              <button key={c} onClick={() => set('color', c)}
                className="w-7 h-7 rounded-full border-2 transition-all"
                style={{ background: c, borderColor: form.color === c ? '#0D0D0D' : 'transparent' }} />
            ))}
          </div>
        </div>
        <div className="flex gap-3 pt-2">
          <button onClick={onClose} className="btn-outline flex-1">Cancel</button>
          <button onClick={handleSave} className="btn-primary flex-1">
            {fund ? 'Update' : 'Add Fund'}
          </button>
        </div>
      </div>
    </Modal>
  )
}

function ETFModal({ etf, onSave, onClose }: {
  etf?: ETF; onSave: (e: ETF) => void; onClose: () => void
}) {
  const [form, setForm] = useState<Omit<ETF, 'id'>>({
    name: etf?.name ?? '',
    symbol: etf?.symbol ?? '',
    category: etf?.category ?? 'Large Cap',
    sipAmount: etf?.sipAmount ?? 1000,
    startDate: etf?.startDate ?? currentYearMonth(),
    expectedReturn: etf?.expectedReturn ?? 12,
    color: etf?.color ?? ASSET_COLORS[2],
  })

  const set = (k: keyof typeof form, v: any) => setForm((p) => ({ ...p, [k]: v }))

  const handleSave = () => {
    if (!form.name.trim() || !form.symbol.trim()) return
    onSave({ ...form, id: etf?.id ?? generateId() })
  }

  return (
    <Modal title={etf ? 'Edit ETF' : 'Add New ETF'} onClose={onClose}>
      <div className="space-y-4">
        <div>
          <label className="text-xs font-medium text-muted mb-1.5 block uppercase tracking-wide">ETF Name</label>
          <input className="input-field" placeholder="e.g. Nippon India ETF Nifty50" value={form.name}
            onChange={(e) => set('name', e.target.value)} />
        </div>
        <div>
          <label className="text-xs font-medium text-muted mb-1.5 block uppercase tracking-wide">Symbol/Ticker</label>
          <input className="input-field" placeholder="e.g. NIPPONETF" value={form.symbol}
            onChange={(e) => set('symbol', e.target.value.toUpperCase())} />
        </div>
        <div className="grid grid-cols-2 gap-3">
          <div>
            <label className="text-xs font-medium text-muted mb-1.5 block uppercase tracking-wide">Category</label>
            <select className="input-field" value={form.category} onChange={(e) => set('category', e.target.value)}>
              {ETF_CATEGORIES.map((c) => <option key={c}>{c}</option>)}
            </select>
          </div>
          <div>
            <label className="text-xs font-medium text-muted mb-1.5 block uppercase tracking-wide">Monthly SIP (₹)</label>
            <input className="input-field" type="number" min={100} step={100} value={form.sipAmount}
              onChange={(e) => set('sipAmount', Number(e.target.value))} />
          </div>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <div>
            <label className="text-xs font-medium text-muted mb-1.5 block uppercase tracking-wide">Start Month</label>
            <input className="input-field" type="month" value={form.startDate}
              onChange={(e) => set('startDate', e.target.value)} />
          </div>
          <div>
            <label className="text-xs font-medium text-muted mb-1.5 block uppercase tracking-wide">Expected CAGR %</label>
            <input className="input-field" type="number" min={1} max={50} step={0.5} value={form.expectedReturn}
              onChange={(e) => set('expectedReturn', Number(e.target.value))} />
          </div>
        </div>
        <div>
          <label className="text-xs font-medium text-muted mb-2 block uppercase tracking-wide">Color</label>
          <div className="flex gap-2 flex-wrap">
            {ASSET_COLORS.map((c) => (
              <button key={c} onClick={() => set('color', c)}
                className="w-7 h-7 rounded-full border-2 transition-all"
                style={{ background: c, borderColor: form.color === c ? '#0D0D0D' : 'transparent' }} />
            ))}
          </div>
        </div>
        <div className="flex gap-3 pt-2">
          <button onClick={onClose} className="btn-outline flex-1">Cancel</button>
          <button onClick={handleSave} className="btn-primary flex-1">
            {etf ? 'Update' : 'Add ETF'}
          </button>
        </div>
      </div>
    </Modal>
  )
}

function StockModal({ stock, onSave, onClose }: {
  stock?: Stock; onSave: (s: Stock) => void; onClose: () => void
}) {
  const [form, setForm] = useState<Omit<Stock, 'id'>>({
    name: stock?.name ?? '',
    symbol: stock?.symbol ?? '',
    category: stock?.category ?? 'Large Cap',
    quantity: stock?.quantity ?? 0,
    buyPrice: stock?.buyPrice ?? 0,
    startDate: stock?.startDate ?? currentYearMonth(),
    expectedReturn: stock?.expectedReturn ?? 15,
    color: stock?.color ?? ASSET_COLORS[4],
  })

  const set = (k: keyof typeof form, v: any) => setForm((p) => ({ ...p, [k]: v }))

  const handleSave = () => {
    if (!form.name.trim() || !form.symbol.trim()) return
    onSave({ ...form, id: stock?.id ?? generateId() })
  }

  return (
    <Modal title={stock ? 'Edit Stock' : 'Add New Stock'} onClose={onClose}>
      <div className="space-y-4">
        <div>
          <label className="text-xs font-medium text-muted mb-1.5 block uppercase tracking-wide">Company Name</label>
          <input className="input-field" placeholder="e.g. Reliance Industries" value={form.name}
            onChange={(e) => set('name', e.target.value)} />
        </div>
        <div>
          <label className="text-xs font-medium text-muted mb-1.5 block uppercase tracking-wide">Symbol/Ticker</label>
          <input className="input-field" placeholder="e.g. RELIANCE" value={form.symbol}
            onChange={(e) => set('symbol', e.target.value.toUpperCase())} />
        </div>
        <div className="grid grid-cols-2 gap-3">
          <div>
            <label className="text-xs font-medium text-muted mb-1.5 block uppercase tracking-wide">Category</label>
            <select className="input-field" value={form.category} onChange={(e) => set('category', e.target.value)}>
              {STOCK_CATEGORIES.map((c) => <option key={c}>{c}</option>)}
            </select>
          </div>
          <div>
            <label className="text-xs font-medium text-muted mb-1.5 block uppercase tracking-wide">Quantity</label>
            <input className="input-field" type="number" min={1} step={1} value={form.quantity}
              onChange={(e) => set('quantity', Number(e.target.value))} />
          </div>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <div>
            <label className="text-xs font-medium text-muted mb-1.5 block uppercase tracking-wide">Buy Price (₹)</label>
            <input className="input-field" type="number" min={1} step={0.01} value={form.buyPrice}
              onChange={(e) => set('buyPrice', Number(e.target.value))} />
          </div>
          <div>
            <label className="text-xs font-medium text-muted mb-1.5 block uppercase tracking-wide">Start Month</label>
            <input className="input-field" type="month" value={form.startDate}
              onChange={(e) => set('startDate', e.target.value)} />
          </div>
        </div>
        <div>
          <label className="text-xs font-medium text-muted mb-1.5 block uppercase tracking-wide">Expected CAGR %</label>
          <input className="input-field" type="number" min={1} max={50} step={0.5} value={form.expectedReturn}
            onChange={(e) => set('expectedReturn', Number(e.target.value))} />
        </div>
        <div>
          <label className="text-xs font-medium text-muted mb-2 block uppercase tracking-wide">Color</label>
          <div className="flex gap-2 flex-wrap">
            {ASSET_COLORS.map((c) => (
              <button key={c} onClick={() => set('color', c)}
                className="w-7 h-7 rounded-full border-2 transition-all"
                style={{ background: c, borderColor: form.color === c ? '#0D0D0D' : 'transparent' }} />
            ))}
          </div>
        </div>
        <div className="flex gap-3 pt-2">
          <button onClick={onClose} className="btn-outline flex-1">Cancel</button>
          <button onClick={handleSave} className="btn-primary flex-1">
            {stock ? 'Update' : 'Add Stock'}
          </button>
        </div>
      </div>
    </Modal>
  )
}

// ─── Add Payment/Investment Modal ─────────────────────────────────────────

function InvestmentModal({
  assetType, funds, etfs, stocks, payment, onSave, onClose
}: {
  assetType: AssetType;
  funds: Fund[];
  etfs: ETF[];
  stocks: Stock[];
  payment?: Payment;
  onSave: (p: Payment) => void;
  onClose: () => void
}) {
  const getAssets = () => {
    switch (assetType) {
      case 'mutual-funds': return funds;
      case 'etfs': return etfs;
      case 'stocks': return stocks;
      default: return [];
    }
  }

  const getAssetName = (asset: any) => {
    if ('symbol' in asset) return `${asset.name} (${asset.symbol})`;
    return asset.name;
  }

  const getDefaultAmount = (asset: any) => {
    if ('sipAmount' in asset) return asset.sipAmount;
    if ('quantity' in asset) return asset.quantity * asset.buyPrice;
    return 1000;
  }

  const [form, setForm] = useState({
    assetId: payment?.assetId ?? (getAssets()[0]?.id ?? ''),
    assetType: payment?.assetType ?? assetType,
    date: payment?.date ?? currentYearMonth(),
    amount: payment?.amount ?? (getAssets()[0] ? getDefaultAmount(getAssets()[0]) : 1000),
    quantity: payment?.quantity ?? '',
    price: payment?.price ?? '',
    notes: payment?.notes ?? '',
  })

  const set = (k: keyof typeof form, v: any) => setForm((p) => ({ ...p, [k]: v }))

  const handleAssetChange = (id: string) => {
    const asset = getAssets().find((x) => x.id === id)
    set('assetId', id)
    if (asset && !payment) {
      if ('sipAmount' in asset) set('amount', asset.sipAmount)
      if ('quantity' in asset && asset.quantity && asset.buyPrice) {
        set('amount', asset.quantity * asset.buyPrice)
        set('quantity', asset.quantity)
        set('price', asset.buyPrice)
      }
    }
  }

  const handleSave = () => {
    if (!form.assetId || !form.date) return
    const p: Payment = {
      id: payment?.id ?? generateId(),
      assetId: form.assetId,
      assetType: form.assetType,
      date: form.date,
      amount: Number(form.amount),
    }
    if (form.quantity) p.quantity = Number(form.quantity)
    if (form.price) p.price = Number(form.price)
    if (form.notes) p.notes = String(form.notes)
    onSave(p)
  }

  const assetTypeLabel = assetType === 'mutual-funds' ? 'Fund' : assetType === 'etfs' ? 'ETF' : 'Stock'

  return (
    <Modal title={payment ? 'Edit Investment' : `Log ${assetTypeLabel} Investment`} onClose={onClose}>
      <div className="space-y-4">
        <div>
          <label className="text-xs font-medium text-muted mb-1.5 block uppercase tracking-wide">{assetTypeLabel}</label>
          <select className="input-field" value={form.assetId} onChange={(e) => handleAssetChange(e.target.value)}>
            {getAssets().map((asset) => <option key={asset.id} value={asset.id}>{getAssetName(asset)}</option>)}
          </select>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <div>
            <label className="text-xs font-medium text-muted mb-1.5 block uppercase tracking-wide">Month</label>
            <input className="input-field" type="month" value={form.date}
              onChange={(e) => set('date', e.target.value)} />
          </div>
          <div>
            <label className="text-xs font-medium text-muted mb-1.5 block uppercase tracking-wide">Amount (₹)</label>
            <input className="input-field" type="number" min={1} step={100} value={form.amount}
              onChange={(e) => set('amount', e.target.value)} />
          </div>
        </div>
        {(assetType === 'stocks' || assetType === 'etfs') && (
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="text-xs font-medium text-muted mb-1.5 block uppercase tracking-wide">Quantity</label>
              <input className="input-field" type="number" min={1} step={1} placeholder="e.g. 10"
                value={form.quantity as string} onChange={(e) => set('quantity', e.target.value)} />
            </div>
            <div>
              <label className="text-xs font-medium text-muted mb-1.5 block uppercase tracking-wide">Price per share (₹)</label>
              <input className="input-field" type="number" step="0.01" placeholder="e.g. 2450.50"
                value={form.price as string} onChange={(e) => set('price', e.target.value)} />
            </div>
          </div>
        )}
        <div>
          <label className="text-xs font-medium text-muted mb-1.5 block uppercase tracking-wide">Notes</label>
          <input className="input-field" placeholder="optional" value={form.notes as string}
            onChange={(e) => set('notes', e.target.value)} />
        </div>
        <div className="flex gap-3 pt-2">
          <button onClick={onClose} className="btn-outline flex-1">Cancel</button>
          <button onClick={handleSave} className="btn-primary flex-1">
            {payment ? 'Update' : 'Log Investment'}
          </button>
        </div>
      </div>
    </Modal>
  )
}

// ─── Projection Calculator ────────────────────────────────────────────────────

function ProjectionModal({ funds, etfs, stocks, payments, onClose }: {
  funds: Fund[]; etfs: ETF[]; stocks: Stock[]; payments: Payment[]; onClose: () => void
}) {
  const [years, setYears] = useState(10)
  const allAssets = [...funds, ...etfs, ...stocks]

  const getMonthlyAmount = (asset: any) => {
    if ('sipAmount' in asset) return asset.sipAmount;
    if ('quantity' in asset && asset.quantity) return (asset.quantity * asset.buyPrice) / 12;
    return 0;
  }

  const data = Array.from({ length: years }, (_, i) => {
    const month = (i + 1) * 12
    const row: any = { year: `Yr ${i + 1}` }
    let totalInvested = 0
    let totalValue = 0

    allAssets.forEach((asset) => {
      const monthlyAmt = getMonthlyAmount(asset)
      const invested = monthlyAmt * month
      const fv = calcFutureValue(monthlyAmt, month, asset.expectedReturn)
      row[asset.name] = Math.round(fv)
      totalInvested += invested
      totalValue += fv
    })

    row['Total Invested'] = Math.round(totalInvested)
    row['Projected Value'] = Math.round(totalValue)
    return row
  })

  const totalMonthly = allAssets.reduce((s, asset) => s + getMonthlyAmount(asset), 0)
  const projectedAtEnd = data[years - 1]?.['Projected Value'] ?? 0

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4" onClick={onClose}>
      <div className="absolute inset-0 bg-ink/40 backdrop-blur-sm" />
      <div className="relative bg-cream rounded-2xl shadow-2xl border border-gold/20 w-full max-w-2xl p-6"
        onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between mb-5">
          <h3 className=" text-lg font-semibold text-ink">Portfolio Projection Calculator</h3>
          <button onClick={onClose} className="text-muted hover:text-ink"><X size={18} /></button>
        </div>

        <div className="flex items-center gap-4 mb-5">
          <span className="text-sm text-muted">Project for</span>
          <div className="flex items-center gap-2 flex-wrap">
            {[3, 5, 10, 15, 20, 25, 30].map((y) => (
              <button key={y} onClick={() => setYears(y)}
                className={`px-3 py-1.5 rounded-lg text-sm font-medium transition-all ${years === y ? 'bg-gold text-ink' : 'bg-paper text-muted hover:text-ink'}`}>
                {y}Y
              </button>
            ))}
          </div>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-3 gap-3 mb-5">
          <div className="stat-card">
            <span className="text-xs text-muted uppercase tracking-wide">Monthly Investment</span>
            <span className=" font-bold text-xl text-emerald">₹{totalMonthly.toLocaleString('en-IN')}</span>
          </div>
          <div className="stat-card">
            <span className="text-xs text-muted uppercase tracking-wide">Total Invested ({years}Y)</span>
            <span className=" font-bold text-xl text-ink">₹{(totalMonthly * years * 12).toLocaleString('en-IN')}</span>
          </div>
          <div className="stat-card">
            <span className="text-xs text-muted uppercase tracking-wide">Projected Value</span>
            <span className=" font-bold text-xl text-gold-dark">₹{Math.round(projectedAtEnd / 100000).toLocaleString('en-IN')}L</span>
          </div>
        </div>

        <ResponsiveContainer width="100%" height={220}>
          <AreaChart data={data}>
            <defs>
              <linearGradient id="gInvested" x1="0" y1="0" x2="0" y2="1">
                <stop offset="5%" stopColor="#8B6914" stopOpacity={0.3} />
                <stop offset="95%" stopColor="#8B6914" stopOpacity={0} />
              </linearGradient>
              <linearGradient id="gValue" x1="0" y1="0" x2="0" y2="1">
                <stop offset="5%" stopColor="#1A5C3A" stopOpacity={0.4} />
                <stop offset="95%" stopColor="#1A5C3A" stopOpacity={0} />
              </linearGradient>
            </defs>
            <XAxis dataKey="year" tick={{ fontSize: 11, fill: '#8A8070' }} />
            <YAxis tickFormatter={(v) => `₹${(v / 100000).toFixed(0)}L`} tick={{ fontSize: 11, fill: '#8A8070' }} />
            <Tooltip formatter={(v: any) => [`₹${Number(v).toLocaleString('en-IN')}`, '']}
              contentStyle={{ background: '#FDF8F0', border: '1px solid #C9A84C33', borderRadius: 12 }} />
            <Area type="monotone" dataKey="Total Invested" stroke="#8B6914" fill="url(#gInvested)" strokeWidth={2} />
            <Area type="monotone" dataKey="Projected Value" stroke="#1A5C3A" fill="url(#gValue)" strokeWidth={2} />
          </AreaChart>
        </ResponsiveContainer>
        <p className="text-xs text-muted text-center mt-2">Projections are estimates based on expected CAGR. Actual returns may vary.</p>
      </div>
    </div>
  )
}

// ─── Main App ─────────────────────────────────────────────────────────────────
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
export default function Home() {
  const [data, setData] = useState<TrackerData>(emptyData())
  const [loaded, setLoaded] = useState(false)
  const [mainTab, setMainTab] = useState<'dashboard' | 'mutual-funds' | 'etfs' | 'stocks'>('dashboard')
  const [subTab, setSubTab] = useState<'list' | 'investments'>('list')

  // Modal states
  const [showFundModal, setShowFundModal] = useState(false)
  const [editFund, setEditFund] = useState<Fund | undefined>()
  const [showETFModal, setShowETFModal] = useState(false)
  const [editETF, setEditETF] = useState<ETF | undefined>()
  const [showStockModal, setShowStockModal] = useState(false)
  const [editStock, setEditStock] = useState<Stock | undefined>()
  const [showInvestmentModal, setShowInvestmentModal] = useState(false)
  const [investmentAssetType, setInvestmentAssetType] = useState<AssetType>('mutual-funds')
  const [editPayment, setEditPayment] = useState<Payment | undefined>()
  const [showProjection, setShowProjection] = useState(false)

  const fileRef = useRef<HTMLInputElement>(null)
  const [isClient, setIsClient] = useState(false)

  // Add this useEffect right after your other useEffects
  useEffect(() => {
    setIsClient(true)
  }, [])

  // Load data from localStorage on mount
  useEffect(() => {
    if (typeof window !== 'undefined') {
      const savedData = localStorage.getItem(STORAGE_KEY)
      if (savedData) {
        try {
          const parsedData = JSON.parse(savedData)
          setData(parsedData)
          setLoaded(true)
        } catch (error) {
          console.error('Failed to parse saved data:', error)
          setLoaded(false)
        }
      } else {
        setLoaded(false)
      }
    }
  }, [])

  // Save data to localStorage whenever it changes
  const updateData = useCallback((newData: TrackerData | ((prev: TrackerData) => TrackerData)) => {
    setData((prevData) => {
      const updatedData = typeof newData === 'function' ? newData(prevData) : newData
      // Save to localStorage (only on client-side)
      if (typeof window !== 'undefined') {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(updatedData))
      }
      return updatedData
    })
  }, [])

  // ── File handling ──
  const handleUpload = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (!file) return
    const reader = new FileReader()
    reader.onload = (ev) => {
      try {
        const parsed = parseExcel(ev.target!.result as ArrayBuffer)
        updateData(parsed)
        setLoaded(true)
      } catch {
        alert('Could not read the Excel file. Make sure it\'s a valid investment-tracker.xlsx file.')
      }
    }
    reader.readAsArrayBuffer(file)
    e.target.value = ''
  }, [updateData])

  const handleExport = () => exportExcel(data)

  const startFresh = () => {
    updateData(emptyData())
    setLoaded(true)
  }



  // ── Funds CRUD ──
  const saveFund = (f: Fund) => {
    updateData((d) => ({
      ...d,
      funds: d.funds.find((x) => x.id === f.id)
        ? d.funds.map((x) => x.id === f.id ? f : x)
        : [...d.funds, f],
    }))
    setShowFundModal(false); setEditFund(undefined)
  }
  const deleteFund = (id: string) => {
    if (!confirm('Delete this fund and all its investments?')) return
    updateData((d) => ({
      ...d,
      funds: d.funds.filter((f) => f.id !== id),
      payments: d.payments.filter((p) => !(p.assetId === id && p.assetType === 'mutual-funds'))
    }))
  }

  // ── ETFs CRUD ──
  const saveETF = (e: ETF) => {
    updateData((d) => ({
      ...d,
      etfs: d.etfs.find((x) => x.id === e.id)
        ? d.etfs.map((x) => x.id === e.id ? e : x)
        : [...d.etfs, e],
    }))
    setShowETFModal(false); setEditETF(undefined)
  }
  const deleteETF = (id: string) => {
    if (!confirm('Delete this ETF and all its investments?')) return
    updateData((d) => ({
      ...d,
      etfs: d.etfs.filter((e) => e.id !== id),
      payments: d.payments.filter((p) => !(p.assetId === id && p.assetType === 'etfs'))
    }))
  }

  // ── Stocks CRUD ──
  const saveStock = (s: Stock) => {
    updateData((d) => ({
      ...d,
      stocks: d.stocks.find((x) => x.id === s.id)
        ? d.stocks.map((x) => x.id === s.id ? s : x)
        : [...d.stocks, s],
    }))
    setShowStockModal(false); setEditStock(undefined)
  }
  const deleteStock = (id: string) => {
    if (!confirm('Delete this stock and all its investments?')) return
    updateData((d) => ({
      ...d,
      stocks: d.stocks.filter((s) => s.id !== id),
      payments: d.payments.filter((p) => !(p.assetId === id && p.assetType === 'stocks'))
    }))
  }

  // ── Payments CRUD ──
  const savePayment = (p: Payment) => {
    updateData((d) => ({
      ...d,
      payments: d.payments.find((x) => x.id === p.id)
        ? d.payments.map((x) => x.id === p.id ? p : x)
        : [...d.payments, p],
    }))
    setShowInvestmentModal(false); setEditPayment(undefined)
  }
  const deletePayment = (id: string) => {
    updateData((d) => ({ ...d, payments: d.payments.filter((p) => p.id !== id) }))
  }

  // ── Derived stats ──
  const totalInvested = calcTotalInvested(data.payments)
  const totalMonthly = [...data.funds, ...data.etfs].reduce((s, asset) => s + ('sipAmount' in asset ? asset.sipAmount : 0), 0)
  const totalStockValue = data.stocks.reduce((s, stock) => s + (stock.quantity * stock.buyPrice), 0)
  const uniqueMonths = [...new Set(data.payments.map((p) => p.date))].length

  // Get investments for current asset type
  const getCurrentInvestments = () => {
    switch (mainTab) {
      case 'mutual-funds': return data.payments.filter(p => p.assetType === 'mutual-funds');
      case 'etfs': return data.payments.filter(p => p.assetType === 'etfs');
      case 'stocks': return data.payments.filter(p => p.assetType === 'stocks');
      default: return [];
    }
  }

  // Get assets for current tab
  const getCurrentAssets = () => {
    switch (mainTab) {
      case 'mutual-funds': return data.funds;
      case 'etfs': return data.etfs;
      case 'stocks': return data.stocks;
      default: return [];
    }
  }

  // Pie chart data for dashboard
  const pieData = [
    ...data.funds.map((f) => ({
      name: f.name,
      value: calcTotalInvested(data.payments, f.id, 'mutual-funds'),
      color: f.color,
      type: 'Mutual Fund'
    })),
    ...data.etfs.map((e) => ({
      name: e.name,
      value: calcTotalInvested(data.payments, e.id, 'etfs'),
      color: e.color,
      type: 'ETF'
    })),
    ...data.stocks.map((s) => ({
      name: s.name,
      value: calcTotalInvested(data.payments, s.id, 'stocks'),
      color: s.color,
      type: 'Stock'
    }))
  ].filter((x) => x.value > 0)

  // Timeline data for dashboard
  const allMonths = [...new Set(data.payments.map((p) => p.date))].sort()
  const timelineData = allMonths.map((m) => {
    const row: any = { month: getMonthLabel(m) }
    const monthPayments = data.payments.filter((p) => p.date === m)
    row.mutualFunds = monthPayments.filter(p => p.assetType === 'mutual-funds').reduce((s, p) => s + p.amount, 0)
    row.etfs = monthPayments.filter(p => p.assetType === 'etfs').reduce((s, p) => s + p.amount, 0)
    row.stocks = monthPayments.filter(p => p.assetType === 'stocks').reduce((s, p) => s + p.amount, 0)
    row.total = monthPayments.reduce((s, p) => s + p.amount, 0)
    return row
  })

  // Function to check if data exists
  const hasData = data.funds.length > 0 || data.etfs.length > 0 || data.stocks.length > 0

  // ── Landing ──
  // Show loading state while checking localStorage on client
  if (!isClient) {
    return (
      <div className="min-h-screen bg-paper flex flex-col items-center justify-center p-6">
        <div className="max-w-lg w-full text-center">
          <div className="w-16 h-16 rounded-2xl bg-emerald flex items-center justify-center mx-auto mb-6 animate-pulse">
            <Logo size={44} />
          </div>
          <div className="h-8 bg-gray-200 rounded w-3/4 mx-auto mb-3 animate-pulse"></div>
          <div className="h-4 bg-gray-200 rounded w-full mx-auto mb-2 animate-pulse"></div>
          <div className="h-4 bg-gray-200 rounded w-2/3 mx-auto mb-10 animate-pulse"></div>
        </div>
      </div>
    )
  }

  // Only render the actual content on client
  if (!loaded) {
    const hasSavedData = localStorage.getItem(STORAGE_KEY)

    return (
      <div className="min-h-screen bg-paper flex flex-col items-center justify-center p-6">
        <div className="max-w-lg w-full text-center">
          <div className="w-16 h-16 rounded-2xl bg-emerald flex items-center justify-center mx-auto mb-6">
            <TrendingUp size={28} className="text-cream" />
          </div>
          <h1 className="text-4xl font-bold text-ink mb-3">Investment Tracker</h1>
          <p className="text-muted text-base mb-2">Track Mutual Funds, ETFs & Stocks in one place.</p>
          <p className="text-muted text-sm mb-10">
            All data lives in your <span className="text-gold-dark font-medium">browser's localStorage</span> — no accounts, no servers, no cookies.
          </p>

          {/* Show Continue button if data exists in localStorage */}
          {hasSavedData && (
            <button className="btn-primary w-full mb-4" onClick={() => setLoaded(true)}>
              <RefreshCw size={14} className="inline mr-2" />Continue with Saved Data
            </button>
          )}

          <div className="card p-6 mb-4 text-left">
            <div className="flex items-start gap-3 mb-4">
              <FileSpreadsheet size={20} className="text-gold mt-0.5 shrink-0" />
              <div>
                <p className="font-medium text-sm text-ink">Load from Excel file</p>
                <p className="text-xs text-muted mt-0.5">Upload your previously saved <code className="bg-paper px-1 rounded text-gold-dark">investment-tracker.xlsx</code></p>
              </div>
            </div>
            <input ref={fileRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={handleUpload} />
            <button className="btn-gold w-full" onClick={() => fileRef.current?.click()}>
              <Upload size={14} className="inline mr-2" />Upload Excel File
            </button>
          </div>

          <div className="relative my-4">
            <div className="absolute inset-0 flex items-center"><div className="w-full border-t border-gold/20" /></div>
            <div className="relative text-center"><span className="bg-paper px-3 text-xs text-muted">or</span></div>
          </div>

          <button className="btn-outline w-full" onClick={startFresh}>
            <Plus size={14} className="inline mr-2" />Start Fresh
          </button>

          <div className="mt-8 flex items-start gap-2 text-xs text-muted bg-cream rounded-xl p-3 text-left">
            <Info size={14} className="shrink-0 mt-0.5 text-gold" />
            <span>Data is automatically saved to your browser. Always download a backup Excel file for safekeeping.</span>
          </div>
        </div>
      </div>
    )
  }

  const currentInvestments = getCurrentInvestments()
  const currentAssets = getCurrentAssets()

  return (
    <div className="min-h-screen bg-paper">
      {/* Header */}
      <header className="bg-cream border-b border-gold/15 sticky top-0 z-30">
        <div className="max-w-6xl mx-auto px-4 py-3 flex items-center justify-between">
          <div className="flex items-center gap-1">
            <div className="w-8 h-8 rounded-xl bg-emerald flex items-center justify-center">
              <Logo />
            </div>
            <span className="font-bold text-ink text-base leading-none">TrackMyFund</span>
          </div>
          <div className="flex items-center gap-2">
            <button className="btn-outline text-xs py-2 px-3 min-w-fit" onClick={() => fileRef.current?.click()}>
              <Upload size={13} className="inline mr-1.5" />Load
            </button>
            <button className="btn-gold text-xs py-2 px-3 min-w-fit" onClick={handleExport}>
              <Download size={13} className="inline mr-1.5" />Save Excel
            </button>
            <input ref={fileRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={handleUpload} />
          </div>
        </div>
        {/* Main Tabs */}
        <div className="max-w-6xl mx-auto px-4 flex gap-1 pb-0 overflow-auto">
          {([
            ['dashboard', 'Dashboard', BarChart3],
            ['mutual-funds', 'Mutual Funds', Wallet],
            ['etfs', 'ETFs', BarChart2],
            ['stocks', 'Stocks', Activity]
          ] as const).map(([id, label, Icon]) => (
            <button key={id} onClick={() => { setMainTab(id); setSubTab('list') }}
              className={`flex items-center min-w-fit gap-1.5 px-2 py-2.5 text-sm font-medium border-b-2 transition-all -mb-px ${mainTab === id ? 'border-gold text-gold-dark' : 'border-transparent text-muted hover:text-ink'}`}>
              <Icon size={14} />{label}
            </button>
          ))}
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-4 py-6">

        {/* ── DASHBOARD ── */}
        {mainTab === 'dashboard' && (
          <div className="space-y-6">
            {/* Stats row */}
            <div className="grid grid-cols-2 md:grid-cols-5 gap-3">
              {[
                { label: 'Total Invested', value: `₹${totalInvested.toLocaleString('en-IN')}`, sub: 'all time', icon: Wallet, color: 'text-emerald' },
                { label: 'Monthly SIP', value: `₹${totalMonthly.toLocaleString('en-IN')}`, sub: 'per month', icon: Calendar, color: 'text-gold-dark' },
                { label: 'Stock Value', value: `₹${totalStockValue.toLocaleString('en-IN')}`, sub: 'lump sum', icon: TrendingUp, color: 'text-blue-600' },
                { label: 'Active Assets', value: (data.funds.length + data.etfs.length + data.stocks.length).toString(), sub: 'total holdings', icon: TrendingUp, color: 'text-ink' },
                { label: 'Months Tracked', value: uniqueMonths.toString(), sub: 'investments logged', icon: BarChart3, color: 'text-ink' },
              ].map((s) => (
                <div key={s.label} className="stat-card">
                  <span className="text-xs text-muted uppercase tracking-wide">{s.label}</span>
                  <span className={` font-bold text-2xl ${s.color}`}>{s.value}</span>
                  <span className="text-xs text-muted">{s.sub}</span>
                </div>
              ))}
            </div>

            {data.funds.length === 0 && data.etfs.length === 0 && data.stocks.length === 0 ? (
              <div className="card p-12 text-center">
                <TrendingUp size={40} className="text-gold/40 mx-auto mb-4" />
                <p className=" text-lg text-ink mb-1">No investments yet</p>
                <p className="text-sm text-muted mb-5">Add mutual funds, ETFs or stocks to get started</p>
                <div className="flex gap-3 justify-center">
                  <button className="btn-outline min-w-fit text-xs" onClick={() => setMainTab('mutual-funds')}>Add Fund</button>
                  <button className="btn-outline min-w-fit text-xs" onClick={() => setMainTab('etfs')}>Add ETF</button>
                  <button className="btn-outline min-w-fit text-xs" onClick={() => setMainTab('stocks')}>Add Stock</button>
                </div>
              </div>
            ) : (
              <div className="grid md:grid-cols-2 gap-5">
                {/* Allocation pie */}
                <div className="card p-5">
                  <h3 className=" font-semibold text-ink mb-4">Portfolio Allocation</h3>
                  {pieData.length === 0 ? <p className="text-sm text-muted text-center py-8">Log investments to see allocation</p> : (
                    <ResponsiveContainer width="100%" height={220}>
                      <PieChart>
                        <Pie data={pieData} cx="50%" cy="50%" outerRadius={80} dataKey="value"
                          label={({ name, percent }) => `${name.split(' ').slice(-1)[0]} ${(percent * 100).toFixed(0)}%`}
                          labelLine={false}>
                          {pieData.map((e, i) => <Cell key={i} fill={e.color} />)}
                        </Pie>
                        <Tooltip formatter={(v: any) => [`₹${Number(v).toLocaleString('en-IN')}`, 'Invested']}
                          contentStyle={{ background: '#FDF8F0', border: '1px solid #C9A84C33', borderRadius: 12 }} />
                      </PieChart>
                    </ResponsiveContainer>
                  )}
                </div>

                {/* Monthly bar chart */}
                <div className="card p-5">
                  <h3 className=" font-semibold text-ink mb-4">Monthly Investment by Type</h3>
                  {timelineData.length === 0 ? <p className="text-sm text-muted text-center py-8">Log investments to see timeline</p> : (
                    <ResponsiveContainer width="100%" height={220}>
                      <BarChart data={timelineData.slice(-12)}>
                        <XAxis dataKey="month" tick={{ fontSize: 10, fill: '#8A8070' }} />
                        <YAxis tickFormatter={(v) => `₹${(v / 1000).toFixed(0)}k`} tick={{ fontSize: 10, fill: '#8A8070' }} />
                        <Tooltip formatter={(v: any) => [`₹${Number(v).toLocaleString('en-IN')}`, '']}
                          contentStyle={{ background: '#FDF8F0', border: '1px solid #C9A84C33', borderRadius: 12 }} />
                        <Bar dataKey="mutualFunds" stackId="a" fill="#C9A84C" name="Mutual Funds" />
                        <Bar dataKey="etfs" stackId="a" fill="#1A5C3A" name="ETFs" />
                        <Bar dataKey="stocks" stackId="a" fill="#1565C0" name="Stocks" />
                      </BarChart>
                    </ResponsiveContainer>
                  )}
                </div>
              </div>
            )}

            {/* Asset summary cards */}
            {[...data.funds, ...data.etfs, ...data.stocks].length > 0 && (
              <div>
                <div className="flex items-center justify-between mb-3">
                  <h3 className=" font-semibold text-ink">Asset Overview</h3>
                  <button className="btn-outline text-xs py-2 px-3" onClick={() => setShowProjection(true)}>
                    <TrendingUp size={13} className="inline mr-1.5" />Projection
                  </button>
                </div>
                <div className="grid md:grid-cols-2 gap-3">
                  {[...data.funds, ...data.etfs, ...data.stocks].slice(0, 6).map((asset) => {
                    const invested = calcTotalInvested(data.payments, asset.id, asset.assetType ||
                      ('symbol' in asset && 'quantity' in asset ? 'stocks' :
                        'symbol' in asset ? 'etfs' : 'mutual-funds') as AssetType)
                    const monthlyAmt = 'sipAmount' in asset ? asset.sipAmount : 0
                    const projected10y = calcFutureValue(monthlyAmt, 120, asset.expectedReturn)
                    return (
                      <div key={asset.id} className="card p-4 flex gap-4 items-start">
                        <div className="w-3 h-full min-h-[60px] rounded-full shrink-0" style={{ background: asset.color }} />
                        <div className="flex-1 min-w-0">
                          <div className="flex items-start justify-between gap-2">
                            <div>
                              <p className="font-medium text-sm text-ink truncate">{asset.name}</p>
                              <span className="text-xs text-muted">{asset.category} · {asset.expectedReturn}% CAGR</span>
                            </div>
                            <div className="text-right shrink-0">
                              <p className="text-xs text-muted">Invested</p>
                              <p className=" font-semibold text-emerald">₹{invested.toLocaleString('en-IN')}</p>
                            </div>
                          </div>
                          {'sipAmount' in asset && (
                            <div className="flex gap-4 mt-2 text-xs text-muted">
                              <span>SIP: <strong className="text-ink">₹{asset.sipAmount.toLocaleString('en-IN')}/mo</strong></span>
                              <span>10Y est: <strong className="text-gold-dark">₹{(projected10y / 100000).toFixed(1)}L</strong></span>
                            </div>
                          )}
                          {'quantity' in asset && (
                            <div className="flex gap-4 mt-2 text-xs text-muted">
                              <span>Qty: <strong className="text-ink">{asset.quantity}</strong></span>
                              <span>Avg Price: <strong className="text-ink">₹{asset.buyPrice.toLocaleString('en-IN')}</strong></span>
                            </div>
                          )}
                        </div>
                      </div>
                    )
                  })}
                </div>
              </div>
            )}

            <div className="card p-4 flex items-center gap-3 bg-emerald/5 border-emerald/20">
              <Info size={16} className="text-emerald shrink-0" />
              <p className="text-xs text-muted">
                Data is automatically saved to your browser. Remember to <button className="text-gold-dark font-medium underline underline-offset-2" onClick={handleExport}>download your Excel file</button> as a backup.
              </p>
            </div>
          </div>
        )}

        {/* ── ASSET MANAGEMENT (Mutual Funds, ETFs, Stocks) ── */}
        {mainTab !== 'dashboard' && (
          <div className="space-y-4">
            {/* Sub tabs */}
            <div className="flex items-center justify-between">
              <div className="flex gap-2 border-b border-gold/20">
                <button
                  onClick={() => setSubTab('list')}
                  className={`px-4 py-2 text-sm font-medium transition-all ${subTab === 'list' ? 'text-gold-dark border-b-2 border-gold' : 'text-muted hover:text-ink'}`}
                >
                  Assets List
                </button>
                <button
                  onClick={() => setSubTab('investments')}
                  className={`px-4 py-2 text-sm font-medium transition-all ${subTab === 'investments' ? 'text-gold-dark border-b-2 border-gold' : 'text-muted hover:text-ink'}`}
                >
                  Investment History
                </button>
              </div>
              <button
                className="btn-primary min-w-fit bg-emerald text-white"
                onClick={() => {
                  if (mainTab === 'mutual-funds') setShowFundModal(true)
                  else if (mainTab === 'etfs') setShowETFModal(true)
                  else setShowStockModal(true)
                }}
              >
                <Plus size={14} className="inline mr-1.5" />
                Add {mainTab === 'mutual-funds' ? 'Fund' : mainTab === 'etfs' ? 'ETF' : 'Stock'}
              </button>
            </div>

            {/* Assets List View */}
            {subTab === 'list' && (
              <>
                {currentAssets.length === 0 ? (
                  <div className="card p-12 text-center">
                    <Wallet size={40} className="text-gold/40 mx-auto mb-4" />
                    <p className=" text-lg text-ink mb-1">No {mainTab === 'mutual-funds' ? 'funds' : mainTab === 'etfs' ? 'ETFs' : 'stocks'} added yet</p>
                    <p className="text-sm text-muted mb-5">Add your first {mainTab === 'mutual-funds' ? 'mutual fund' : mainTab === 'etfs' ? 'ETF' : 'stock'} to start tracking</p>
                    <button className="btn-primary bg-emerald text-white" onClick={() => {
                      if (mainTab === 'mutual-funds') setShowFundModal(true)
                      else if (mainTab === 'etfs') setShowETFModal(true)
                      else setShowStockModal(true)
                    }}>
                      Add First {mainTab === 'mutual-funds' ? 'Fund' : mainTab === 'etfs' ? 'ETF' : 'Stock'}
                    </button>
                  </div>
                ) : (
                  <div className="space-y-3">
                    {currentAssets.map((asset) => {
                      const invested = calcTotalInvested(data.payments, asset.id, mainTab)
                      const investments = data.payments.filter((p) => p.assetId === asset.id && p.assetType === mainTab)
                      const monthlyAmt = 'sipAmount' in asset ? asset.sipAmount : 0
                      return (
                        <div key={asset.id} className="card p-5">
                          <div className="flex items-start gap-4">
                            <div className="w-10 h-10 rounded-xl flex items-center justify-center shrink-0" style={{ background: asset.color + '22' }}>
                              <TrendingUp size={18} style={{ color: asset.color }} />
                            </div>
                            <div className="flex-1 min-w-0">
                              <div className="flex items-start justify-between gap-3">
                                <div>
                                  <h3 className="font-semibold text-ink">{asset.name}</h3>
                                  {'symbol' in asset && <p className="text-xs text-muted mt-0.5">Symbol: {asset.symbol}</p>}
                                  <p className="text-xs text-muted mt-0.5">{asset.category} · Since {asset.startDate} · {asset.expectedReturn}% expected CAGR</p>
                                </div>
                                <div className="flex gap-1 shrink-0">
                                  <button
                                    onClick={() => {
                                      setInvestmentAssetType(mainTab)
                                      setEditPayment(undefined)
                                      setShowInvestmentModal(true)
                                    }}
                                    className="p-2 text-emerald hover:text-emerald-dark rounded-lg hover:bg-paper transition-all"
                                    title="Add Investment"
                                  >
                                    <Plus size={14} />
                                  </button>
                                  <button onClick={() => {
                                    if (mainTab === 'mutual-funds') { setEditFund(asset as Fund); setShowFundModal(true) }
                                    else if (mainTab === 'etfs') { setEditETF(asset as ETF); setShowETFModal(true) }
                                    else { setEditStock(asset as Stock); setShowStockModal(true) }
                                  }} className="p-2 text-muted hover:text-ink rounded-lg hover:bg-paper transition-all">
                                    <Edit3 size={14} />
                                  </button>
                                  <button onClick={() => {
                                    if (mainTab === 'mutual-funds') deleteFund(asset.id)
                                    else if (mainTab === 'etfs') deleteETF(asset.id)
                                    else deleteStock(asset.id)
                                  }} className="p-2 text-muted hover:text-red-500 rounded-lg hover:bg-paper transition-all">
                                    <Trash2 size={14} />
                                  </button>
                                </div>
                              </div>
                              <div className="grid grid-cols-1 lg:grid-cols-3 gap-3 mt-3">
                                {'sipAmount' in asset && (
                                  <div className="bg-paper rounded-xl p-3">
                                    <p className="text-xs text-muted">Monthly SIP</p>
                                    <p className=" font-bold text-emerald">₹{asset.sipAmount.toLocaleString('en-IN')}</p>
                                  </div>
                                )}
                                {'quantity' in asset && (
                                  <div className="bg-paper rounded-xl p-3">
                                    <p className="text-xs text-muted">Quantity</p>
                                    <p className=" font-bold text-ink">{asset.quantity}</p>
                                  </div>
                                )}
                                {'quantity' in asset && (
                                  <div className="bg-paper rounded-xl p-3">
                                    <p className="text-xs text-muted">Buy Price</p>
                                    <p className=" font-bold text-ink">₹{asset.buyPrice.toLocaleString('en-IN')}</p>
                                  </div>
                                )}
                                <div className="bg-paper rounded-xl p-3">
                                  <p className="text-xs text-muted">Total Invested</p>
                                  <p className=" font-bold text-ink">₹{invested.toLocaleString('en-IN')}</p>
                                </div>
                                <div className="bg-paper rounded-xl p-3">
                                  <p className="text-xs text-muted">Investments</p>
                                  <p className=" font-bold text-ink">{investments.length} entries</p>
                                </div>
                              </div>
                            </div>
                          </div>
                        </div>
                      )
                    })}
                  </div>
                )}
              </>
            )}

            {/* Investment History View */}
            {subTab === 'investments' && (
              <>
                {currentAssets.length === 0 ? (
                  <div className="card p-12 text-center">
                    <Calendar size={40} className="text-gold/40 mx-auto mb-4" />
                    <p className=" text-lg text-ink mb-1">No assets added yet</p>
                    <p className="text-sm text-muted mb-5">Add assets before logging investments</p>
                    <button className="btn-primary" onClick={() => {
                      if (mainTab === 'mutual-funds') setShowFundModal(true)
                      else if (mainTab === 'etfs') setShowETFModal(true)
                      else setShowStockModal(true)
                    }}>
                      Add First {mainTab === 'mutual-funds' ? 'Fund' : mainTab === 'etfs' ? 'ETF' : 'Stock'}
                    </button>
                  </div>
                ) : currentInvestments.length === 0 ? (
                  <div className="card p-12 text-center">
                    <Calendar size={40} className="text-gold/40 mx-auto mb-4" />
                    <p className=" text-lg text-ink mb-1">No investments logged yet</p>
                    <p className="text-sm text-muted mb-5">Start recording your {mainTab === 'mutual-funds' ? 'SIP payments' : mainTab === 'etfs' ? 'ETF purchases' : 'stock purchases'}</p>
                    <button className="btn-primary" onClick={() => {
                      setInvestmentAssetType(mainTab)
                      setEditPayment(undefined)
                      setShowInvestmentModal(true)
                    }}>
                      Log First Investment
                    </button>
                  </div>
                ) : (
                  <div className="card overflow-hidden">
                    <div className="overflow-x-auto scrollbar-thin">
                      <table className="w-full text-sm">
                        <thead>
                          <tr className="border-b border-gold/15 bg-paper/50">
                            <th className="text-left px-4 py-3 text-xs font-medium text-muted uppercase tracking-wide">Month</th>
                            <th className="text-left px-4 py-3 text-xs font-medium text-muted uppercase tracking-wide">Name</th>
                            <th className="text-right px-4 py-3 text-xs font-medium text-muted uppercase tracking-wide">Amount</th>
                            {(mainTab === 'etfs' || mainTab === 'stocks') && (
                              <>
                                <th className="text-right px-4 py-3 text-xs font-medium text-muted uppercase tracking-wide">Quantity</th>
                                <th className="text-right px-4 py-3 text-xs font-medium text-muted uppercase tracking-wide">Price</th>
                              </>
                            )}
                            <th className="text-left px-4 py-3 text-xs font-medium text-muted uppercase tracking-wide">Notes</th>
                            <th className="px-4 py-3" />
                          </tr>
                        </thead>
                        <tbody>
                          {[...currentInvestments].sort((a, b) => b.date.localeCompare(a.date)).map((p) => {
                            const asset = currentAssets.find((a) => a.id === p.assetId)
                            return (
                              <tr key={p.id} className="border-b border-gold/10 hover:bg-paper/40 transition-colors">
                                <td className="px-4 py-3 font-mono text-xs text-muted">{getMonthLabel(p.date)}</td>
                                <td className="px-4 py-3">
                                  <div className="flex items-center gap-2">
                                    <div className="w-2 h-2 rounded-full" style={{ background: asset?.color ?? '#ccc' }} />
                                    <span className="truncate max-w-[180px]">{asset?.name ?? 'Unknown'}</span>
                                  </div>
                                </td>
                                <td className="px-4 py-3 text-right font-medium text-emerald">₹{p.amount.toLocaleString('en-IN')}</td>
                                {(mainTab === 'etfs' || mainTab === 'stocks') && (
                                  <>
                                    <td className="px-4 py-3 text-right text-muted">{p.quantity || '—'}</td>
                                    <td className="px-4 py-3 text-right text-muted">{p.price ? `₹${p.price}` : '—'}</td>
                                  </>
                                )}
                                <td className="px-4 py-3 text-muted text-xs truncate max-w-[120px]">{p.notes || '—'}</td>
                                <td className="px-4 py-3">
                                  <div className="flex gap-1 justify-end">
                                    <button onClick={() => {
                                      setEditPayment(p)
                                      setInvestmentAssetType(mainTab)
                                      setShowInvestmentModal(true)
                                    }} className="p-1.5 text-muted hover:text-ink rounded-lg hover:bg-paper transition-all">
                                      <Edit3 size={13} />
                                    </button>
                                    <button onClick={() => deletePayment(p.id)} className="p-1.5 text-muted hover:text-red-500 rounded-lg hover:bg-paper transition-all">
                                      <Trash2 size={13} />
                                    </button>
                                  </div>
                                </td>
                              </tr>
                            )
                          })}
                        </tbody>
                        <tfoot>
                          <tr className="bg-paper/50">
                            <td colSpan={2} className="px-4 py-3 text-xs font-medium text-muted uppercase tracking-wide">Total</td>
                            <td className="px-4 py-3 text-right font-bold text-emerald">₹{currentInvestments.reduce((s, p) => s + p.amount, 0).toLocaleString('en-IN')}</td>
                            {(mainTab === 'etfs' || mainTab === 'stocks') && (
                              <>
                                <td colSpan={2} />
                              </>
                            )}
                            <td colSpan={2} />
                          </tr>
                        </tfoot>
                      </table>
                    </div>
                  </div>
                )}
              </>
            )}
          </div>
        )}
      </main>

      {/* Modals */}
      {showFundModal && (
        <FundModal fund={editFund} onSave={saveFund} onClose={() => { setShowFundModal(false); setEditFund(undefined) }} />
      )}
      {showETFModal && (
        <ETFModal etf={editETF} onSave={saveETF} onClose={() => { setShowETFModal(false); setEditETF(undefined) }} />
      )}
      {showStockModal && (
        <StockModal stock={editStock} onSave={saveStock} onClose={() => { setShowStockModal(false); setEditStock(undefined) }} />
      )}
      {showInvestmentModal && (
        <InvestmentModal
          assetType={investmentAssetType}
          funds={data.funds}
          etfs={data.etfs}
          stocks={data.stocks}
          payment={editPayment}
          onSave={savePayment}
          onClose={() => { setShowInvestmentModal(false); setEditPayment(undefined) }}
        />
      )}
      {showProjection && (
        <ProjectionModal funds={data.funds} etfs={data.etfs} stocks={data.stocks} payments={data.payments} onClose={() => setShowProjection(false)} />
      )}
      <span className='flex justify-center text-muted text-sm pb-4'>Copyright © 2026 Sudeep Teja.</span>
    </div>
  )
}