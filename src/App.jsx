import { useState, useCallback, useRef } from "react"
import * as XLSX from "xlsx"
import {
  BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer,
  RadarChart, Radar, PolarGrid, PolarAngleAxis, Cell,
} from "recharts"

/* ─── PR Score Model ─────────────────────────────────────── */
const WEIGHTS = { cvr: 0.35, ctr: 0.25, bounce: 0.20, cart: 0.15, search: 0.05 }

function normalize(arr, inverse = false) {
  const mn = Math.min(...arr)
  const mx = Math.max(...arr)
  return arr.map(v => {
    const s = mx === mn ? 50 : ((v - mn) / (mx - mn)) * 100
    return inverse ? 100 - s : s
  })
}

function calcPR(products) {
  if (!products.length) return []
  const ctrs    = normalize(products.map(p => p.ctr))
  const cvrs    = normalize(products.map(p => p.cvr))
  const bounces = normalize(products.map(p => p.bounce), true)
  const carts   = normalize(products.map(p => p.cart))
  const searches= normalize(products.map(p => p.search))
  return products.map((p, i) => ({
    ...p,
    pr: Math.round(
      cvrs[i]    * WEIGHTS.cvr    +
      ctrs[i]    * WEIGHTS.ctr    +
      bounces[i] * WEIGHTS.bounce +
      carts[i]   * WEIGHTS.cart   +
      searches[i]* WEIGHTS.search
    ),
    ctrN: ctrs[i], cvrN: cvrs[i], bounceN: 100 - bounces[i],
    cartN: carts[i],
  }))
}

function getDirection(p) {
  const dirs = []
  if (p.ctr < 2.0)                    dirs.push("主圖")
  if (p.bounce > 20)                  dirs.push("商品說明")
  if (p.cart > 12 && p.cvr < 4)       dirs.push("價格")
  if (p.cvr === 0 && p.ctr > 0)       dirs.push("主圖")
  if (!dirs.length)                    dirs.push("維持")
  return [...new Set(dirs)]
}

/* ─── Parse xlsx ──────────────────────────────────────────── */
function parseXlsx(buffer) {
  const wb = XLSX.read(buffer, { type: "array" })
  const products = []

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName]
    const rows = XLSX.utils.sheet_to_json(ws, { defval: "" })
    if (!rows.length) continue

    const keys = Object.keys(rows[0])
    const has = k => keys.some(kk => kk.includes(k))
    if (!has("商品ID") || !has("曝光")) continue

    for (const row of rows) {
      const id      = String(row["商品ID"] || "")
      const name    = row["商品名稱"] || row["商品名"] || ""
      const specId  = row["商品規格ID"] || row["規格ID"] || ""

      // Only parent rows (no variant spec ID, or spec = "-")
      if (specId && specId !== "-" && specId !== "") continue
      if (!id || !name || name.includes("無法獲取")) continue

      const getNum = (label) => {
        const k = keys.find(kk => kk.includes(label))
        if (!k) return 0
        const v = String(row[k] || "0").replace(/[,%]/g, "")
        return parseFloat(v) || 0
      }

      products.push({
        id,
        name: name.length > 26 ? name.slice(0, 26) + "…" : name,
        fullName: name,
        ctr:    getNum("點擊率"),
        cvr:    getNum("轉換率 (全部"),
        bounce: getNum("跳出率"),
        cart:   getNum("加入購物車轉換率"),
        search: getNum("搜尋點擊"),
        sales:  getNum("銷售額(全部"),
        orders: getNum("全部訂單") || getNum("訂單"),
        impressions: getNum("曝光次數"),
      })
    }
  }

  return products
}

/* ─── Helpers ─────────────────────────────────────────────── */
const prColor = pr =>
  pr >= 65 ? "#00e599" : pr >= 45 ? "#ffb547" : "#ff4d6d"

const dirStyle = {
  "主圖":   { bg: "rgba(0,212,255,0.12)", color: "#00d4ff" },
  "商品說明":{ bg: "rgba(0,229,153,0.12)", color: "#00e599" },
  "價格":   { bg: "rgba(255,181,71,0.12)", color: "#ffb547" },
  "維持":   { bg: "rgba(90,98,130,0.15)",  color: "#5a6282" },
}

const fmt = n => Math.round(n).toLocaleString("zh-TW")
const fmtPct = n => n.toFixed(1) + "%"

/* ─── Demo data (preloaded) ──────────────────────────────── */
const DEMO = [
  { id:"44807163242", name:"100%純豆腐貓砂 3/6/10入",   fullName:"100%純豆腐貓砂 3/6/10入", ctr:2.37, cvr:5.37, bounce:15.91, cart:17.61, search:108, sales:6454,  orders:13, impressions:10213 },
  { id:"44760145523", name:"豆腐貓砂 長期9折方案",      fullName:"豆腐貓砂 長期9折方案",     ctr:2.25, cvr:7.95, bounce:12.50, cart:17.86, search:35,  sales:3202,  orders:7,  impressions:3907  },
  { id:"44559729248", name:"豆腐貓砂 天然大豆",         fullName:"豆腐貓砂 天然大豆",        ctr:1.78, cvr:7.50, bounce:35.76, cart:13.94, search:62,  sales:3886,  orders:9,  impressions:6743  },
  { id:"57457399697", name:"豆腐砂/木薯砂 混搭",        fullName:"豆腐砂/木薯砂 混搭",       ctr:2.16, cvr:4.58, bounce:14.68, cart:19.27, search:63,  sales:3350,  orders:7,  impressions:7073  },
  { id:"53809704079", name:"豆腐貓砂 整箱8入",          fullName:"豆腐貓砂 整箱8入",         ctr:2.21, cvr:2.86, bounce:13.43, cart:17.91, search:27,  sales:2535,  orders:3,  impressions:4760  },
  { id:"54657274107", name:"混合貓砂 豆腐木薯雙效",     fullName:"混合貓砂 豆腐木薯雙效",    ctr:1.63, cvr:4.88, bounce:20.00, cart:5.71,  search:23,  sales:1668,  orders:2,  impressions:2521  },
  { id:"56060051090", name:"木薯貓砂",                  fullName:"木薯貓砂",                 ctr:2.35, cvr:2.48, bounce:20.21, cart:11.70, search:70,  sales:971,   orders:3,  impressions:5155  },
  { id:"47660150063", name:"全系列貓砂體驗包",          fullName:"全系列貓砂體驗包",         ctr:1.42, cvr:3.64, bounce:24.44, cart:15.56, search:22,  sales:908,   orders:2,  impressions:3864  },
  { id:"46510067048", name:"礦石混合貓砂",              fullName:"礦石混合貓砂",             ctr:1.49, cvr:4.84, bounce:18.52, cart:7.41,  search:44,  sales:835,   orders:3,  impressions:4155  },
  { id:"53107441845", name:"膠囊式漏砂板貓砂盆",        fullName:"膠囊式漏砂板貓砂盆",       ctr:6.80, cvr:3.01, bounce:21.90, cart:10.48, search:60,  sales:808,   orders:4,  impressions:1955  },
  { id:"52509708131", name:"豆腐貓砂 2包體驗組",        fullName:"豆腐貓砂 2包體驗組",       ctr:1.29, cvr:2.78, bounce:13.79, cart:6.90,  search:16,  sales:598,   orders:1,  impressions:2796  },
  { id:"48657298011", name:"100%木薯貓砂 3/4/8入",      fullName:"100%木薯貓砂 3/4/8入",    ctr:1.65, cvr:0,    bounce:0,     cart:0,     search:12,  sales:0,     orders:0,  impressions:1215  },
  { id:"51457446999", name:"SPA寵物按摩器",             fullName:"SPA寵物按摩器",            ctr:2.47, cvr:0,    bounce:66.67, cart:0,     search:5,   sales:0,     orders:0,  impressions:365   },
]

/* ─── Components ──────────────────────────────────────────── */
function Tag({ label }) {
  const s = dirStyle[label] || dirStyle["維持"]
  return (
    <span style={{
      background: s.bg, color: s.color,
      border: `1px solid ${s.color}33`,
      padding: "2px 8px", borderRadius: 20,
      fontSize: 11, fontWeight: 500,
      display: "inline-block", marginRight: 4,
    }}>{label}</span>
  )
}

function PRBar({ pr }) {
  const c = prColor(pr)
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
      <div style={{
        flex: 1, height: 6, background: "#1c2035",
        borderRadius: 3, overflow: "hidden",
      }}>
        <div style={{ width: `${pr}%`, height: "100%", background: c, borderRadius: 3, transition: "width .4s" }} />
      </div>
      <span style={{ fontSize: 13, fontWeight: 500, color: c, minWidth: 26, textAlign: "right" }}>{pr}</span>
    </div>
  )
}

function MetricBadge({ label, value, good, warn }) {
  const num = parseFloat(value)
  const color = num >= good ? "#00e599" : num >= warn ? "#ffb547" : "#ff4d6d"
  return (
    <div style={{ display: "inline-block" }}>
      <span style={{ fontSize: 11, color: "#5a6282" }}>{label} </span>
      <span style={{ fontSize: 12, fontWeight: 500, color }}>{value}</span>
    </div>
  )
}

function DropZone({ onFile }) {
  const [drag, setDrag] = useState(false)
  const inputRef = useRef()

  const handle = useCallback(file => {
    if (!file) return
    const r = new FileReader()
    r.onload = e => onFile(new Uint8Array(e.target.result))
    r.readAsArrayBuffer(file)
  }, [onFile])

  return (
    <div
      onClick={() => inputRef.current.click()}
      onDragOver={e => { e.preventDefault(); setDrag(true) }}
      onDragLeave={() => setDrag(false)}
      onDrop={e => { e.preventDefault(); setDrag(false); handle(e.dataTransfer.files[0]) }}
      style={{
        border: `1.5px dashed ${drag ? "var(--cyan)" : "var(--border-md)"}`,
        borderRadius: "var(--radius-lg)",
        padding: "28px 20px",
        textAlign: "center",
        cursor: "pointer",
        transition: "border-color .2s",
        background: drag ? "rgba(0,212,255,0.04)" : "transparent",
        marginBottom: 24,
      }}
    >
      <div style={{ fontSize: 28, marginBottom: 8 }}>📂</div>
      <div style={{ fontSize: 13, color: "#8892b0" }}>
        拖曳或點擊上傳 <span style={{ color: "var(--cyan)" }}>parentskudetail_*.xlsx</span>
      </div>
      <div style={{ fontSize: 11, color: "var(--muted)", marginTop: 4 }}>
        蝦皮賣家中心 → 數據 → 商品分析 → 匯出
      </div>
      <input ref={inputRef} type="file" accept=".xlsx" style={{ display: "none" }}
        onChange={e => handle(e.target.files[0])} />
    </div>
  )
}

function WeightLegend() {
  const items = [
    { label: "轉換率 CVR", w: 35, color: "#00d4ff" },
    { label: "點擊率 CTR",  w: 25, color: "#00d4ff" },
    { label: "跳出率（負向）", w: 20, color: "#ff4d6d" },
    { label: "加入購物車率", w: 15, color: "#00e599" },
    { label: "搜尋點擊",    w: 5,  color: "#5a6282" },
  ]
  return (
    <div style={{
      background: "var(--card)", border: "0.5px solid var(--border)",
      borderRadius: "var(--radius-lg)", padding: "14px 16px", marginBottom: 20,
    }}>
      <div style={{ fontSize: 11, color: "var(--muted)", letterSpacing: ".06em", marginBottom: 10 }}>演算法權重模型</div>
      {items.map(it => (
        <div key={it.label} style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 7 }}>
          <span style={{ fontSize: 12, color: "var(--text)", minWidth: 120 }}>{it.label}</span>
          <div style={{ flex: 1, height: 5, background: "#1c2035", borderRadius: 3, overflow: "hidden" }}>
            <div style={{ width: `${it.w * 2}%`, height: "100%", background: it.color, borderRadius: 3 }} />
          </div>
          <span style={{ fontSize: 12, fontWeight: 500, color: it.color, minWidth: 28, textAlign: "right" }}>{it.w}%</span>
        </div>
      ))}
      <div style={{ fontSize: 11, color: "var(--muted)", marginTop: 8, lineHeight: 1.5 }}>
        依據 Cloud Ecommerce（2026/02）、Shopdora（2025/11）等公開研究整合。CVR 在所有來源均列為最高權重指標；跳出率為負向扣分項。
      </div>
    </div>
  )
}

function DirectionGuide() {
  const items = [
    { key: "主圖", color: "#00d4ff", icon: "🖼️", title: "主圖優化", desc: "CTR < 2% 代表在搜尋頁失去點擊。強化主圖標語（「可沖馬桶」「低粉塵」）、白底清晰照、差異化文字標語。" },
    { key: "商品說明", color: "#00e599", icon: "📄", title: "商品說明優化", desc: "跳出率 > 20% 代表顧客進頁後快速離開。優化首圖組（使用場景）、規格清晰度、評價展示區。頁面開頭 3 秒決定去留。" },
    { key: "價格", color: "#ffb547", icon: "🏷️", title: "價格/組合優化", desc: "購物車率高但 CVR 低，代表定價摩擦。測試限時折扣、搭購組合、滿額優惠，降低最終結帳的心理阻力。" },
  ]
  return (
    <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 12, marginBottom: 24 }}>
      {items.map(it => (
        <div key={it.key} style={{
          background: "var(--card)", border: "0.5px solid var(--border)",
          borderRadius: "var(--radius-lg)", padding: "14px 14px",
        }}>
          <div style={{ fontSize: 14, marginBottom: 6 }}>{it.icon}
            <span style={{ fontSize: 12, fontWeight: 500, color: it.color, marginLeft: 6 }}>{it.title}</span>
          </div>
          <div style={{ fontSize: 12, color: "#8892b0", lineHeight: 1.6 }}>{it.desc}</div>
        </div>
      ))}
    </div>
  )
}

/* ─── Radar detail panel ─────────────────────────────────── */
function DetailPanel({ product, onClose }) {
  if (!product) return null
  const radarData = [
    { metric: "CTR",    value: product.ctrN  },
    { metric: "CVR",    value: product.cvrN  },
    { metric: "不跳出",  value: 100 - product.bounceN },
    { metric: "購物車",  value: product.cartN },
  ]
  const dirs = getDirection(product)
  return (
    <div style={{
      position: "fixed", inset: 0, background: "rgba(8,11,18,0.85)",
      display: "flex", alignItems: "center", justifyContent: "center", zIndex: 100,
    }} onClick={onClose}>
      <div onClick={e => e.stopPropagation()} style={{
        background: "var(--surface)", border: "0.5px solid var(--border-md)",
        borderRadius: 16, padding: "24px 28px", width: 480, maxWidth: "92vw",
      }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 16 }}>
          <div>
            <div style={{ fontSize: 13, fontWeight: 500, color: "var(--text)", marginBottom: 4 }}>{product.fullName}</div>
            <div style={{ fontSize: 11, color: "var(--muted)" }}>ID: {product.id}</div>
          </div>
          <button onClick={onClose} style={{ color: "var(--muted)", fontSize: 18, lineHeight: 1 }}>✕</button>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 8, marginBottom: 16 }}>
          {[
            { label: "銷售額", val: `NT$${fmt(product.sales)}` },
            { label: "訂單數", val: product.orders },
            { label: "PR 分數", val: product.pr, color: prColor(product.pr) },
            { label: "CTR",  val: fmtPct(product.ctr) },
            { label: "CVR",  val: fmtPct(product.cvr) },
            { label: "跳出率", val: fmtPct(product.bounce) },
          ].map(m => (
            <div key={m.label} style={{ background: "var(--card)", borderRadius: 8, padding: "10px 12px" }}>
              <div style={{ fontSize: 11, color: "var(--muted)", marginBottom: 2 }}>{m.label}</div>
              <div style={{ fontSize: 16, fontWeight: 500, color: m.color || "var(--text)" }}>{m.val}</div>
            </div>
          ))}
        </div>

        <ResponsiveContainer width="100%" height={180}>
          <RadarChart data={radarData}>
            <PolarGrid stroke="rgba(255,255,255,0.07)" />
            <PolarAngleAxis dataKey="metric" tick={{ fill: "#5a6282", fontSize: 11 }} />
            <Radar dataKey="value" stroke="#00d4ff" fill="#00d4ff" fillOpacity={0.18} />
          </RadarChart>
        </ResponsiveContainer>

        <div style={{ marginTop: 12 }}>
          <div style={{ fontSize: 11, color: "var(--muted)", marginBottom: 6 }}>改善方向</div>
          <div>{dirs.map(d => <Tag key={d} label={d} />)}</div>
          {dirs.includes("主圖") && <div style={{ fontSize: 12, color: "#8892b0", marginTop: 6 }}>CTR {fmtPct(product.ctr)} — 建議更換主圖，強化搜尋頁面競爭力。</div>}
          {dirs.includes("商品說明") && <div style={{ fontSize: 12, color: "#8892b0", marginTop: 4 }}>跳出率 {fmtPct(product.bounce)} — 商品頁內容需優化，減少顧客離開。</div>}
          {dirs.includes("價格") && <div style={{ fontSize: 12, color: "#8892b0", marginTop: 4 }}>購物車率 {fmtPct(product.cart)} vs CVR {fmtPct(product.cvr)} — 結帳前流失明顯，建議測試定價或組合。</div>}
        </div>
      </div>
    </div>
  )
}

/* ─── Main App ────────────────────────────────────────────── */
export default function App() {
  const [rawProducts, setRawProducts] = useState(DEMO)
  const [selected, setSelected] = useState(null)
  const [toast, setToast] = useState(null)
  const [tab, setTab] = useState("table") // "table" | "chart" | "guide"

  const showToast = (msg, ok = true) => {
    setToast({ msg, ok })
    setTimeout(() => setToast(null), 3000)
  }

  const handleFile = useCallback(buffer => {
    try {
      const parsed = parseXlsx(buffer)
      if (!parsed.length) { showToast("找不到商品資料，請確認匯出格式", false); return }
      setRawProducts(parsed)
      showToast(`成功匯入 ${parsed.length} 個商品`)
    } catch (e) {
      showToast("解析失敗：" + e.message, false)
    }
  }, [])

  const products = calcPR(rawProducts).sort((a, b) => b.pr - a.pr)
  const totalSales = products.reduce((s, p) => s + p.sales, 0)
  const avgPR = Math.round(products.reduce((s, p) => s + p.pr, 0) / products.length)

  const tabs = [
    { id: "table",  label: "PR 排行" },
    { id: "chart",  label: "指標圖表" },
    { id: "guide",  label: "改善說明" },
  ]

  return (
    <div style={{ minHeight: "100vh", padding: "20px 24px", maxWidth: 900, margin: "0 auto" }}>

      {/* Header */}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-end", marginBottom: 20 }}>
        <div>
          <div style={{ fontFamily: "var(--font-mono)", fontSize: 11, color: "var(--cyan)", letterSpacing: ".1em", marginBottom: 4 }}>SENMAO · 蝦皮商品分析</div>
          <div style={{ fontSize: 20, fontWeight: 500, color: "var(--text)" }}>PR 品質分數儀表板</div>
        </div>
        <div style={{ fontFamily: "var(--font-mono)", fontSize: 11, color: "var(--muted)" }}>
          {products.length} 商品 · 平均 PR {avgPR}
        </div>
      </div>

      {/* KPI row */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 10, marginBottom: 20 }}>
        {[
          { label: "總銷售額", val: `NT$${fmt(totalSales)}`, color: "var(--cyan)" },
          { label: "有銷售商品", val: `${products.filter(p=>p.sales>0).length} / ${products.length}`, color: "var(--text)" },
          { label: "平均 CVR", val: fmtPct(products.reduce((s,p)=>s+p.cvr,0)/products.length), color: "var(--green)" },
          { label: "平均跳出率", val: fmtPct(products.filter(p=>p.bounce>0).reduce((s,p,_,a)=>s+p.bounce/a.length,0)), color: "var(--amber)" },
        ].map(k => (
          <div key={k.label} style={{ background: "var(--card)", border: "0.5px solid var(--border)", borderRadius: "var(--radius)", padding: "12px 14px" }}>
            <div style={{ fontSize: 11, color: "var(--muted)", marginBottom: 4 }}>{k.label}</div>
            <div style={{ fontSize: 18, fontWeight: 500, color: k.color }}>{k.val}</div>
          </div>
        ))}
      </div>

      {/* Upload */}
      <DropZone onFile={handleFile} />

      {/* Tabs */}
      <div style={{ display: "flex", gap: 2, marginBottom: 16, borderBottom: "0.5px solid var(--border)" }}>
        {tabs.map(t => (
          <button key={t.id} onClick={() => setTab(t.id)} style={{
            padding: "8px 16px", fontSize: 13,
            color: tab === t.id ? "var(--cyan)" : "var(--muted)",
            borderBottom: tab === t.id ? "1.5px solid var(--cyan)" : "1.5px solid transparent",
            marginBottom: -1, transition: "color .15s",
          }}>{t.label}</button>
        ))}
      </div>

      {/* Tab: Table */}
      {tab === "table" && (
        <div style={{ background: "var(--card)", border: "0.5px solid var(--border)", borderRadius: "var(--radius-lg)", overflow: "hidden" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", tableLayout: "fixed" }}>
            <thead>
              <tr style={{ borderBottom: "0.5px solid var(--border)" }}>
                {["#","商品","PR 分數","CVR","CTR","跳出率","購物車率","改善方向"].map((h,i) => (
                  <th key={h} style={{
                    padding: "10px 10px", textAlign: i === 0 ? "center" : "left",
                    fontSize: 11, fontWeight: 500, color: "var(--muted)",
                    letterSpacing: ".04em",
                    width: ["32px","28%","14%","7%","7%","7%","8%","auto"][i],
                  }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {products.map((p, i) => {
                const dirs = getDirection(p)
                return (
                  <tr key={p.id}
                    onClick={() => setSelected(p)}
                    style={{ borderBottom: "0.5px solid var(--border)", cursor: "pointer", transition: "background .1s" }}
                    onMouseEnter={e => e.currentTarget.style.background = "#1a1f35"}
                    onMouseLeave={e => e.currentTarget.style.background = "transparent"}
                  >
                    <td style={{ padding: "10px", textAlign: "center", fontSize: 11, color: "var(--muted)" }}>{i + 1}</td>
                    <td style={{ padding: "10px", fontSize: 12, lineHeight: 1.4, color: "var(--text)" }}>{p.name}</td>
                    <td style={{ padding: "10px" }}><PRBar pr={p.pr} /></td>
                    <td style={{ padding: "10px", fontSize: 12, color: p.cvr >= 6 ? "#00e599" : p.cvr >= 3 ? "var(--text)" : "#ff4d6d" }}>{fmtPct(p.cvr)}</td>
                    <td style={{ padding: "10px", fontSize: 12, color: p.ctr >= 2 ? "var(--text)" : "#ffb547" }}>{fmtPct(p.ctr)}</td>
                    <td style={{ padding: "10px", fontSize: 12, color: p.bounce > 25 ? "#ff4d6d" : p.bounce > 15 ? "#ffb547" : "var(--text)" }}>{fmtPct(p.bounce)}</td>
                    <td style={{ padding: "10px", fontSize: 12 }}>{fmtPct(p.cart)}</td>
                    <td style={{ padding: "10px" }}>{dirs.map(d => <Tag key={d} label={d} />)}</td>
                  </tr>
                )
              })}
            </tbody>
          </table>
          <div style={{ padding: "8px 14px", fontSize: 11, color: "var(--muted)" }}>點擊任一列查看詳細雷達圖</div>
        </div>
      )}

      {/* Tab: Chart */}
      {tab === "chart" && (
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
          <div style={{ background: "var(--card)", border: "0.5px solid var(--border)", borderRadius: "var(--radius-lg)", padding: "16px" }}>
            <div style={{ fontSize: 11, color: "var(--muted)", marginBottom: 12 }}>PR 分數分布</div>
            <ResponsiveContainer width="100%" height={280}>
              <BarChart data={products} layout="vertical" margin={{ left: 0, right: 16 }}>
                <XAxis type="number" domain={[0, 100]} tick={{ fill: "#5a6282", fontSize: 10 }} />
                <YAxis type="category" dataKey="name" width={110} tick={{ fill: "#8892b0", fontSize: 10 }} />
                <Tooltip
                  contentStyle={{ background: "#1c2035", border: "0.5px solid #2a3050", borderRadius: 8, fontSize: 12 }}
                  formatter={v => [v, "PR 分數"]}
                />
                <Bar dataKey="pr" radius={[0, 3, 3, 0]}>
                  {products.map(p => <Cell key={p.id} fill={prColor(p.pr)} />)}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
          <div style={{ background: "var(--card)", border: "0.5px solid var(--border)", borderRadius: "var(--radius-lg)", padding: "16px" }}>
            <div style={{ fontSize: 11, color: "var(--muted)", marginBottom: 12 }}>CVR vs CTR（前10商品）</div>
            <ResponsiveContainer width="100%" height={280}>
              <BarChart data={products.slice(0, 10)} margin={{ bottom: 30 }}>
                <XAxis dataKey="name" tick={{ fill: "#5a6282", fontSize: 9 }} angle={-30} textAnchor="end" interval={0} />
                <YAxis tick={{ fill: "#5a6282", fontSize: 10 }} tickFormatter={v => v + "%"} />
                <Tooltip
                  contentStyle={{ background: "#1c2035", border: "0.5px solid #2a3050", borderRadius: 8, fontSize: 12 }}
                  formatter={(v, n) => [fmtPct(v), n]}
                />
                <Bar dataKey="cvr" name="CVR" fill="#00d4ff" radius={[3,3,0,0]} />
                <Bar dataKey="ctr" name="CTR" fill="#00e599" radius={[3,3,0,0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      )}

      {/* Tab: Guide */}
      {tab === "guide" && (
        <div>
          <WeightLegend />
          <DirectionGuide />
          <div style={{ background: "var(--card)", border: "0.5px solid var(--border)", borderRadius: "var(--radius-lg)", padding: "16px", fontSize: 12, color: "#8892b0", lineHeight: 1.8 }}>
            <div style={{ fontWeight: 500, color: "var(--text)", marginBottom: 8 }}>如何使用這個儀表板</div>
            <ol style={{ paddingLeft: 16 }}>
              <li>前往蝦皮賣家中心 → <strong style={{color:"var(--text)"}}>數據 → 商品分析 → 商品概覽</strong>，選擇日期範圍後點「匯出」</li>
              <li>將下載的 <code style={{background:"#1c2035",padding:"1px 5px",borderRadius:3}}>parentskudetail_*.xlsx</code> 拖曳到上方的上傳區域</li>
              <li>系統會自動計算 PR 分數並更新排行表</li>
              <li>點擊任一商品列，查看雷達圖與具體改善建議</li>
              <li>依據改善方向欄位，優先處理標記「主圖」「商品說明」「價格」的商品</li>
            </ol>
          </div>
        </div>
      )}

      {/* Detail modal */}
      {selected && <DetailPanel product={selected} onClose={() => setSelected(null)} />}

      {/* Toast */}
      {toast && (
        <div style={{
          position: "fixed", bottom: 20, left: "50%", transform: "translateX(-50%)",
          background: toast.ok ? "#0f2820" : "#2d0f18",
          border: `0.5px solid ${toast.ok ? "#00e599" : "#ff4d6d"}`,
          color: toast.ok ? "#00e599" : "#ff4d6d",
          padding: "8px 18px", borderRadius: 20, fontSize: 13, fontWeight: 500,
          zIndex: 200,
        }}>{toast.msg}</div>
      )}
    </div>
  )
}
