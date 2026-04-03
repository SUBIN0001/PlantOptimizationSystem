import { useState, useRef, useMemo } from "react";
import * as XLSX from "xlsx";
import logo from "./assets/ashok-leyland-logo.jpg";
// ─────────────────────────────────────────────────────────────────────────────
// CONSTANTS
// ─────────────────────────────────────────────────────────────────────────────

const API = "https://plant-optimization-system-qj11.vercel.app" || "http://localhost:8000/api";

const STEPS = [
  { id: "data", label: "Data Input", sub: "Load Excel / CSV", icon: "⬆" },
  { id: "processing", label: "Processing", sub: "Clean & normalise", icon: "⚙" },
  { id: "constraints", label: "Constraints", sub: "Filter locations", icon: "▽" },
  { id: "optimization", label: "Optimization", sub: "MCDM scoring", icon: "▦" },
  { id: "simulation", label: "Simulation", sub: "Monte Carlo", icon: "◎" },
  { id: "export", label: "Export", sub: "Excel dashboard", icon: "⬇" },
];

const CRITERIA = [
  { key: "vendorBase", label: "Vendor Base", unit: "count", isCost: false },
  { key: "manpowerAvailability", label: "Manpower Availability", unit: "count", isCost: false },
  { key: "capex", label: "CAPEX", unit: "Cr", isCost: true },
  { key: "govtNorms", label: "Govt. Norms", unit: "score", isCost: false },
  { key: "logisticsCost", label: "Logistics Cost", unit: "km", isCost: true },
  { key: "economiesOfScale", label: "Economies of Scale", unit: "score", isCost: false },
];

const DEFAULT_MATRIX = [
  [1, 2, 3, 3, 2, 4],
  [0.5, 1, 2, 2, 2, 3],
  [0.333, 0.5, 1, 1, 0.5, 2],
  [0.333, 0.5, 1, 1, 0.5, 2],
  [0.5, 0.5, 2, 2, 1, 2],
  [0.25, 0.333, 0.5, 0.5, 0.5, 1],
];

// Sub-attribute sections for the upload preview table
const SUB_SECTIONS = [
  {
    key: "location", label: "Location", cols: [
      { key: "name", label: "Location" },
      { key: "region", label: "Region" },
      { key: "state", label: "State" },
      { key: "raw.industrialPark", label: "Industrial Park" },
    ]
  },
  {
    key: "vendor", label: "Vendor Base", score: "vendorBase", cols: [
      { key: "raw.acmaCluster", label: "ACMA Cluster" },
      { key: "raw.acmaUnits", label: "ACMA Units" },
      { key: "raw.tier1Vendors", label: "Tier-1 Vendors" },
      { key: "raw.tier2Vendors", label: "Tier-2 Vendors" },
      { key: "raw.steelSuppliers", label: "Steel Suppliers" },
      { key: "raw.vendorEcosystem", label: "Ecosystem Rating" },
      { key: "raw.keyOEMs", label: "Key OEMs" },
      { key: "vendorBase", label: "★ Score", isScore: true },
    ]
  },
  {
    key: "manpower", label: "Manpower", score: "manpowerAvailability", cols: [
      { key: "raw.enggColleges", label: "Engg Colleges" },
      { key: "raw.itiInstitutes", label: "ITI Institutes" },
      { key: "raw.itiGraduates", label: "ITI Grads (000s)" },
      { key: "raw.enggGraduates", label: "Engg Grads (000s)" },
      { key: "raw.skilledLabourRating", label: "Skill Rating" },
      { key: "raw.wageSkilled", label: "Wage Skilled (₹)" },
      { key: "raw.wageSemiSkilled", label: "Wage Semi (₹)" },
      { key: "raw.attritionRate", label: "Attrition %" },
      { key: "manpowerAvailability", label: "★ Score", isScore: true },
    ]
  },
  {
    key: "capex", label: "CAPEX", score: "capex", cols: [
      { key: "raw.landCost", label: "Land Cost (₹Cr/Ac)" },
      { key: "raw.availableLand", label: "Available Land (Ac)" },
      { key: "raw.constructionIndex", label: "Const. Index" },
      { key: "raw.powerCapex", label: "Power Capex (₹Cr)" },
      { key: "raw.waterCapex", label: "Water Capex (₹Cr)" },
      { key: "raw.totalCapex", label: "★ Total CAPEX (₹Cr)", isScore: true },
    ]
  },
  {
    key: "govt", label: "Govt / Norms", score: "govtNorms", cols: [
      { key: "raw.industrialPolicy", label: "Policy" },
      { key: "raw.capitalSubsidy", label: "Capital Subsidy %" },
      { key: "raw.sgstExemption", label: "SGST Exempt (yrs)" },
      { key: "raw.stampDuty", label: "Stamp Duty" },
      { key: "raw.powerTariff", label: "Power Tariff (₹/kWh)" },
      { key: "raw.elecDutyExemption", label: "Elec Duty Exempt" },
      { key: "raw.envClearanceEase", label: "Env Clear (1-10)" },
      { key: "raw.approvalDays", label: "Approval Days" },
      { key: "raw.sezNimz", label: "SEZ/NIMZ" },
      { key: "raw.dfcAccessGovt", label: "DFC Access" },
      { key: "govtNorms", label: "★ Score", isScore: true },
    ]
  },
  {
    key: "logistics", label: "Logistics", score: "logisticsCost", cols: [
      { key: "raw.nearestPort", label: "Nearest Port" },
      { key: "raw.distanceToPort", label: "Dist Port (km)" },
      { key: "raw.roadConnectivity", label: "Road Conn (1-10)" },
      { key: "raw.railConnectivity", label: "Rail Conn (1-10)" },
      { key: "raw.dfcLogistics", label: "DFC Access" },
      { key: "raw.distanceKeyMarket", label: "Dist Market (km)" },
      { key: "raw.keyMarketCity", label: "Market City" },
      { key: "raw.inboundFreight", label: "Inbound Freight" },
      { key: "raw.outboundFreight", label: "Outbound Freight" },
      { key: "raw.annualLogisticsCost", label: "★ Annual Cost (₹Cr)", isScore: true },
    ]
  },
  {
    key: "economies", label: "Economies of Scale", score: "economiesOfScale", cols: [
      { key: "raw.clusterMaturity", label: "Cluster Maturity" },
      { key: "raw.existingCVOEMs", label: "CV OEMs Nearby" },
      { key: "raw.supplierPark", label: "Supplier Park" },
      { key: "raw.exportHub", label: "Export Hub" },
      { key: "raw.marketDemandIndex", label: "Demand Index" },
      { key: "raw.clusterBenefitScore", label: "Cluster Score" },
      { key: "economiesOfScale", label: "★ Score", isScore: true },
    ]
  },
];

const DEFAULT_CONSTRAINTS = CRITERIA.map((c) => ({
  key: c.key,
  label: c.label,
  operator: c.isCost ? "lte" : "gte",
  value: 0,
  enabled: false,
  isCost: c.isCost,   // preserved so Excel constraint sheet can display direction
}));

// ─────────────────────────────────────────────────────────────────────────────
// UTILITY
// ─────────────────────────────────────────────────────────────────────────────

const fmt = (n, d = 3) => (typeof n === "number" ? n.toFixed(d) : n);
const pct = (n) => `${(n * 100).toFixed(1)}%`;
const scoreColor = (s) => {
  if (s >= 0.7) return "#aa9273";
  if (s >= 0.4) return "#c7b093";
  return "rgba(226,232,240,0.6)";
};

// ─────────────────────────────────────────────────────────────────────────────
// SMALL COMPONENTS
// ─────────────────────────────────────────────────────────────────────────────

function Badge({ rank }) {
  const colors = ["#aa9273", "#c7b093", "#ff8a5c"];
  const bg = rank <= 3 ? colors[rank - 1] : "rgba(30,41,59,0.8)";
  return (
    <span style={{
      background: bg, color: rank <= 3 ? "#fff" : "#94a3b8",
      borderRadius: 6, padding: "2px 8px", fontSize: 11, fontWeight: 700,
    }}>#{rank}</span>
  );
}

function ScoreBar({ score }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
      <div style={{
        flex: 1, height: 6, background: "rgba(30,41,59,0.6)", borderRadius: 3, overflow: "hidden"
      }}>
        <div style={{
          width: pct(score), height: "100%",
          background: `linear-gradient(90deg, ${scoreColor(score)}, ${scoreColor(score)}aa)`,
          borderRadius: 3, transition: "width .6s ease"
        }} />
      </div>
      <span style={{ fontSize: 12, fontWeight: 700, color: scoreColor(score), minWidth: 38 }}>
        {fmt(score)}
      </span>
    </div>
  );
}

function Chip({ label, color = "rgba(30,41,59,0.8)" }) {
  return (
    <span style={{
      background: color, color: "#e2e8f0", borderRadius: 20,
      padding: "2px 10px", fontSize: 11, fontWeight: 600
    }}>{label}</span>
  );
}

function Card({ children, style = {}, glow = false }) {
  return (
    <div style={{
      background: "rgba(6, 16, 40, 0.7)",
      border: `1px solid ${glow ? "rgba(0,200,255,0.25)" : "rgba(0,180,255,0.12)"}`,
      borderRadius: 16,
      padding: 24,
      backdropFilter: "blur(12px)",
      boxShadow: glow
        ? "0 0 30px rgba(0,180,255,0.1), inset 0 0 30px rgba(0,180,255,0.03)"
        : "0 4px 24px rgba(0,0,0,0.3)",
      ...style
    }}>{children}</div>
  );
}

function SectionTitle({ children }) {
  return (
    <h3 style={{
      fontSize: 11, fontWeight: 700, letterSpacing: 2.5,
      textTransform: "uppercase", color: "rgba(100,180,255,0.7)", margin: "0 0 16px",
      fontFamily: "'DM Mono', monospace"
    }}>{children}</h3>
  );
}

function Spinner() {
  return (
    <div style={{ textAlign: "center", padding: 60 }}>
      <div style={{
        width: 48, height: 48,
        border: "2px solid rgba(0,180,255,0.1)",
        borderTop: "2px solid rgba(0,200,255,0.8)",
        borderRadius: "50%",
        animation: "spin 0.8s linear infinite",
        margin: "0 auto 16px",
        boxShadow: "0 0 20px rgba(0,200,255,0.3)"
      }} />
      <p style={{ color: "rgba(100,180,255,0.6)", fontSize: 13, fontFamily: "'DM Mono', monospace", letterSpacing: 1 }}>
        PROCESSING…
      </p>
    </div>
  );
}

function GlowButton({ onClick, children, style = {} }) {
  return (
    <button
      onClick={onClick}
      style={{
        background: "linear-gradient(135deg, rgba(0,160,255,0.2), rgba(0,200,255,0.1))",
        color: "#60d0ff",
        border: "1px solid rgba(0,180,255,0.4)",
        borderRadius: 10,
        padding: "12px 32px",
        fontSize: 12,
        fontWeight: 700,
        cursor: "pointer",
        letterSpacing: 2,
        textTransform: "uppercase",
        fontFamily: "'DM Mono', monospace",
        transition: "all 0.25s ease",
        boxShadow: "0 0 20px rgba(0,180,255,0.15), inset 0 0 20px rgba(0,180,255,0.05)",
        ...style,
      }}
      onMouseEnter={e => {
        e.currentTarget.style.boxShadow = "0 0 35px rgba(0,200,255,0.35), inset 0 0 20px rgba(0,200,255,0.1)";
        e.currentTarget.style.borderColor = "rgba(0,210,255,0.7)";
        e.currentTarget.style.color = "#a0e8ff";
      }}
      onMouseLeave={e => {
        e.currentTarget.style.boxShadow = "0 0 20px rgba(0,180,255,0.15), inset 0 0 20px rgba(0,180,255,0.05)";
        e.currentTarget.style.borderColor = "rgba(0,180,255,0.4)";
        e.currentTarget.style.color = "#60d0ff";
      }}
    >
      {children}
    </button>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// STEP 1 — DATA INPUT
// ─────────────────────────────────────────────────────────────────────────────

function DataInput({ onNext, savedPreview, onPreviewChange }) {
  const [drag, setDrag] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const fileRef = useRef();

  // Use lifted state so preview survives back-navigation
  const preview = savedPreview;
  const setPreview = onPreviewChange;

  const handle = async (file) => {
    if (!file) return;
    setLoading(true); setError(null);
    const fd = new FormData();
    fd.append("file", file);
    try {
      const r = await fetch(`${API}/api/upload`, { method: "POST", body: fd });
      const d = await r.json();
      if (d.error) { setError(d.error); return; }
      setPreview(d);
    } catch (e) {
      setError("Cannot reach backend. Make sure uvicorn is running on port 8000.");
    } finally { setLoading(false); }
  };

  return (
    <div style={{ maxWidth: preview ? 1200 : 860, margin: "0 auto", transition: "max-width 0.4s ease" }}>
      <SectionTitle>Dataset Upload</SectionTitle>

      <div
        onDragOver={(e) => { e.preventDefault(); setDrag(true); }}
        onDragLeave={() => setDrag(false)}
        onDrop={(e) => { e.preventDefault(); setDrag(false); handle(e.dataTransfer.files[0]); }}
        onClick={() => fileRef.current.click()}
        style={{
          border: `2px dashed ${drag ? "rgba(0,200,255,0.7)" : "rgba(0,160,255,0.3)"}`,
          borderRadius: 20, padding: "56px 48px", textAlign: "center",
          cursor: "pointer", transition: "all .3s ease",
          background: drag ? "rgba(0,160,255,0.08)" : "rgba(0,30,80,0.4)",
          backdropFilter: "blur(8px)", marginBottom: 24,
          boxShadow: drag ? "0 0 40px rgba(0,200,255,0.2), inset 0 0 40px rgba(0,200,255,0.05)" : "0 0 0px transparent",
          position: "relative", overflow: "hidden",
        }}
      >
        {[
          { top: 12, left: 12, borderTop: "2px solid rgba(0,200,255,0.5)", borderLeft: "2px solid rgba(0,200,255,0.5)" },
          { top: 12, right: 12, borderTop: "2px solid rgba(0,200,255,0.5)", borderRight: "2px solid rgba(0,200,255,0.5)" },
          { bottom: 12, left: 12, borderBottom: "2px solid rgba(0,200,255,0.5)", borderLeft: "2px solid rgba(0,200,255,0.5)" },
          { bottom: 12, right: 12, borderBottom: "2px solid rgba(0,200,255,0.5)", borderRight: "2px solid rgba(0,200,255,0.5)" },
        ].map((s, i) => (
          <div key={i} style={{ position: "absolute", width: 20, height: 20, ...s }} />
        ))}

        <input ref={fileRef} type="file" accept=".csv,.xlsx,.xls" style={{ display: "none" }}
          onChange={(e) => handle(e.target.files[0])} />

        <div style={{ fontSize: 52, marginBottom: 20, filter: "drop-shadow(0 0 12px rgba(255,180,60,0.5))" }}>📂</div>
        <p style={{ color: "rgba(180,220,255,0.9)", margin: "0 0 8px", fontSize: 16, fontWeight: 500 }}>
          Drop your <strong style={{ color: "#60d0ff" }}>CSV or Excel</strong> dataset here, or click to browse
        </p>
        <p style={{ color: "rgba(100,150,200,0.6)", margin: 0, fontSize: 12, fontFamily: "'DM Mono', monospace" }}>
          Accepted formats: .csv, .xls, .xlsx
        </p>
        <div style={{ marginTop: 24 }}>
          <div style={{
            display: "inline-flex", alignItems: "center", gap: 8,
            background: "linear-gradient(135deg, rgba(0,160,255,0.2), rgba(0,100,200,0.1))",
            border: "1px solid rgba(0,180,255,0.4)", borderRadius: 8, padding: "10px 24px",
            color: "#60d0ff", fontSize: 13, fontWeight: 700, letterSpacing: 1,
            boxShadow: "0 0 20px rgba(0,180,255,0.2)",
          }}>
            📁 Browse Files
          </div>
        </div>
      </div>

      {loading && <Spinner />}
      {error && (
        <div style={{
          color: "#ff6b6b", background: "rgba(255,50,50,0.08)",
          border: "1px solid rgba(255,50,50,0.2)",
          borderRadius: 10, padding: 14, fontSize: 13,
          fontFamily: "'DM Mono', monospace"
        }}>{error}</div>
      )}

      {preview && !loading && (
        <PreviewTable locations={preview.locations} count={preview.count} onNext={() => onNext(preview.locations)} />
      )}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// UPLOAD PREVIEW — tabbed multi-section attribute table
// ─────────────────────────────────────────────────────────────────────────────

function getNestedVal(obj, key) {
  if (!key.includes('.')) return obj[key];
  const [top, rest] = key.split('.', 2);
  const sub = obj[top];
  return sub && typeof sub === 'object' ? sub[rest] : undefined;
}

function PreviewTable({ locations, count, onNext }) {
  const [activeTab, setActiveTab] = useState(0);
  const sec = SUB_SECTIONS[activeTab];

  const fmtCell = (val) => {
    if (val === undefined || val === null || val === '') return '—';
    if (typeof val === 'number') return val.toFixed(2);
    return String(val);
  };

  return (
    <>
      {/* Section tabs */}
      <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap', marginBottom: 12 }}>
        {SUB_SECTIONS.map((s, i) => (
          <button key={s.key} onClick={() => setActiveTab(i)} style={{
            background: i === activeTab ? 'rgba(0,150,255,0.2)' : 'rgba(0,30,80,0.5)',
            border: i === activeTab ? '1px solid rgba(0,200,255,0.6)' : '1px solid rgba(0,100,200,0.2)',
            color: i === activeTab ? '#60d0ff' : 'rgba(100,150,200,0.6)',
            borderRadius: 8, padding: '6px 14px', fontSize: 11, cursor: 'pointer',
            fontFamily: "'DM Mono', monospace", letterSpacing: 0.8, fontWeight: 700,
            boxShadow: i === activeTab ? '0 0 12px rgba(0,180,255,0.2)' : 'none',
            transition: 'all 0.2s',
          }}>{s.label}</button>
        ))}
        <div style={{ marginLeft: 'auto' }}>
          <Chip label={`${count} Locations detected`} color="rgba(0,100,60,0.6)" />
        </div>
      </div>

      {/* Table */}
      <Card style={{ marginBottom: 16, padding: 0 }} glow>
        <div style={{ overflowX: 'auto', maxHeight: 420, overflowY: 'auto' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12, minWidth: 600 }}>
            <thead style={{ position: 'sticky', top: 0, zIndex: 2 }}>
              <tr>
                {/* Always pin Location name in first column */}
                {activeTab !== 0 && (
                  <th style={{
                    textAlign: 'left', padding: '8px 12px', whiteSpace: 'nowrap',
                    background: 'rgba(10,20,60,0.97)', color: 'rgba(100,180,255,0.7)',
                    borderBottom: '1px solid rgba(0,150,255,0.2)',
                    fontFamily: "'DM Mono', monospace", fontSize: 10, letterSpacing: 1,
                    position: 'sticky', left: 0, zIndex: 3,
                  }}>Location</th>
                )}
                {sec.cols.map(c => (
                  <th key={c.key} style={{
                    textAlign: c.isScore ? 'center' : 'left',
                    padding: '8px 12px', whiteSpace: 'nowrap',
                    background: c.isScore ? 'rgba(0,60,30,0.8)' : 'rgba(10,20,60,0.97)',
                    color: c.isScore ? '#4ade80' : 'rgba(100,180,255,0.7)',
                    borderBottom: '1px solid rgba(0,150,255,0.2)',
                    fontFamily: "'DM Mono', monospace", fontSize: 10, letterSpacing: 1,
                  }}>{c.label}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {locations.map((loc, i) => (
                <tr key={i} style={{
                  borderBottom: '1px solid rgba(0,100,200,0.08)',
                  background: i % 2 === 0 ? 'transparent' : 'rgba(0,20,60,0.2)',
                }}>
                  {activeTab !== 0 && (
                    <td style={{
                      padding: '7px 12px', color: '#e2e8f0', fontWeight: 600,
                      whiteSpace: 'nowrap', position: 'sticky', left: 0,
                      background: i % 2 === 0 ? 'rgba(2,11,26,0.97)' : 'rgba(5,18,46,0.97)',
                      zIndex: 1,
                    }}>{loc.name}</td>
                  )}
                  {sec.cols.map(c => {
                    const val = getNestedVal(loc, c.key);
                    return (
                      <td key={c.key} style={{
                        padding: '7px 12px',
                        color: c.isScore ? '#4ade80' : 'rgba(150,200,255,0.75)',
                        textAlign: c.isScore ? 'center' : 'left',
                        fontWeight: c.isScore ? 700 : 400,
                        fontFamily: typeof val === 'number' ? "'DM Mono', monospace" : 'inherit',
                        whiteSpace: 'nowrap',
                      }}>{fmtCell(val)}</td>
                    );
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        {count > locations.length && (
          <p style={{ color: 'rgba(100,150,200,0.5)', fontSize: 12, margin: '8px 16px', fontFamily: "'DM Mono', monospace" }}>
            Showing all {count} rows
          </p>
        )}
      </Card>

      <GlowButton onClick={onNext}>Proceed to Processing →</GlowButton>
    </>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// STEP 2 — PROCESSING
// ─────────────────────────────────────────────────────────────────────────────

function Processing({ locations, onNext }) {
  const checks = [
    { label: "Removed duplicates", status: true },
    { label: "Parsed numeric units (Cr, km)", status: true },
    { label: "Mapped categorical values", status: true },
    { label: "Min-Max normalisation", status: true },
    { label: "Null / NaN values", status: locations.some(l => Object.values(l).some(v => v == null)) },
  ];

  return (
    <div style={{ maxWidth: 700, margin: "0 auto" }}>
      <SectionTitle>Data Cleaning & Normalisation</SectionTitle>
      <Card style={{ marginBottom: 24 }} glow>
        {checks.map((c, i) => (
          <div key={i} style={{
            display: "flex", alignItems: "center", gap: 12,
            padding: "12px 0", borderBottom: i < checks.length - 1 ? "1px solid rgba(0,150,255,0.1)" : "none"
          }}>
            <span style={{ fontSize: 16 }}>{c.status ? "✅" : "⚠️"}</span>
            <span style={{ color: c.status ? "rgba(180,210,255,0.8)" : "#f5a623", fontSize: 14 }}>{c.label}</span>
            <span style={{ marginLeft: "auto", color: c.status ? "#00c896" : "#f5a623", fontSize: 11, fontWeight: 700, fontFamily: "'DM Mono', monospace", letterSpacing: 1 }}>
              {c.status ? "OK" : "WARNING"}
            </span>
          </div>
        ))}
      </Card>

      <Card style={{ marginBottom: 24 }}>
        <SectionTitle>Summary Statistics</SectionTitle>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
            <thead>
              <tr>
                {["Criterion", "Min", "Max", "Mean", "Std"].map(h => (
                  <th key={h} style={{
                    textAlign: "left", padding: "6px 10px",
                    color: "rgba(100,180,255,0.6)",
                    borderBottom: "1px solid rgba(0,150,255,0.15)",
                    fontFamily: "'DM Mono', monospace", fontSize: 10, letterSpacing: 1
                  }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {CRITERIA.map(c => {
                const vals = locations.map(l => parseFloat(l[c.key]) || 0);
                const mean = vals.reduce((a, b) => a + b, 0) / vals.length;
                const std = Math.sqrt(vals.map(v => (v - mean) ** 2).reduce((a, b) => a + b, 0) / vals.length);
                return (
                  <tr key={c.key} style={{ borderBottom: "1px solid rgba(0,100,200,0.08)" }}>
                    <td style={{ padding: "6px 10px", color: "#e2e8f0", fontWeight: 600, fontSize: 13 }}>{c.label}</td>
                    <td style={{ padding: "6px 10px", color: "rgba(100,150,200,0.6)", fontFamily: "'DM Mono', monospace" }}>{fmt(Math.min(...vals), 1)}</td>
                    <td style={{ padding: "6px 10px", color: "rgba(100,150,200,0.6)", fontFamily: "'DM Mono', monospace" }}>{fmt(Math.max(...vals), 1)}</td>
                    <td style={{ padding: "6px 10px", color: "rgba(150,200,255,0.8)", fontFamily: "'DM Mono', monospace" }}>{fmt(mean, 1)}</td>
                    <td style={{ padding: "6px 10px", color: "rgba(150,200,255,0.8)", fontFamily: "'DM Mono', monospace" }}>{fmt(std, 1)}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </Card>

      <GlowButton onClick={onNext}>Set Constraints →</GlowButton>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// STEP 3 — CONSTRAINTS
// ─────────────────────────────────────────────────────────────────────────────

function Constraints({ constraints, onChange, onNext, locations = [] }) {
  // ── Derive unique region + state values as labelled entries ──
  //    { value, label, type } so chips can show type badges.
  //    Filter in Optimization matches l.region or l.state against entry.value.
  const availableRegions = useMemo(() => {
    const seen = new Set();
    const entries = [];
    locations.forEach(l => {
      if (l.region && !seen.has("r::" + l.region)) {
        seen.add("r::" + l.region);
        entries.push({ value: l.region, label: l.region, type: "region" });
      }
      if (l.state && !seen.has("s::" + l.state)) {
        seen.add("s::" + l.state);
        entries.push({ value: l.state, label: l.state, type: "state" });
      }
    });
    return entries.sort((a, b) => a.label.localeCompare(b.label));
  }, [locations]);

  const [regionFilterEnabled, setRegionFilterEnabled] = useState(false);
  const [selectedRegions, setSelectedRegions] = useState([]);

  const toggleRegion = (value) => {
    setSelectedRegions(prev =>
      prev.includes(value) ? prev.filter(r => r !== value) : [...prev, value]
    );
  };

  // Forward both region filter state and step transition to parent
  const handleNext = () => {
    onNext({ regionFilterEnabled, selectedRegions });
  };

  const sharedInputStyle = {
    background: "rgba(0,30,80,0.8)",
    border: "1px solid rgba(0,150,255,0.25)",
    color: "#60d0ff",
    borderRadius: 8,
    padding: "6px 8px",
    fontSize: 13,
    outline: "none",
  };

  return (
    <div style={{ maxWidth: 700, margin: "0 auto" }}>
      <SectionTitle>Feasibility Constraints</SectionTitle>
      <p style={{ color: "rgba(100,150,200,0.6)", fontSize: 13, marginBottom: 24 }}>
        Enable constraints to eliminate infeasible locations before MCDM scoring.
      </p>

      {/* ── Region / State Filter Box ── */}
      <Card style={{ marginBottom: 16 }} glow>
        {/* Header row — identical 4-column grid as constraint rows */}
        <div style={{
          display: "grid",
          gridTemplateColumns: "1fr 120px 100px 80px",
          gap: 12,
          alignItems: "center",
          padding: "14px 0",
          borderBottom: regionFilterEnabled ? "1px solid rgba(0,150,255,0.1)" : "none",
          opacity: regionFilterEnabled ? 1 : 0.45,
          transition: "opacity 0.2s",
        }}>
          <span style={{ color: "#e2e8f0", fontSize: 14, fontWeight: 600 }}>
            Region / State
          </span>

          {/* Operator-column placeholder */}
          <div style={{
            ...sharedInputStyle,
            color: "rgba(100,150,200,0.45)",
            pointerEvents: "none",
            textAlign: "center",
          }}>
            is in
          </div>

          {/* Value-column: live selection count */}
          <div style={{
            ...sharedInputStyle,
            textAlign: "center",
            color: selectedRegions.length > 0 ? "#60d0ff" : "rgba(100,150,200,0.45)",
          }}>
            {selectedRegions.length > 0 ? `${selectedRegions.length} selected` : "all"}
          </div>

          {/* Enable toggle — clearing selection when turned off */}
          <label style={{ display: "flex", alignItems: "center", gap: 6, cursor: "pointer" }}>
            <input
              type="checkbox"
              checked={regionFilterEnabled}
              onChange={e => {
                setRegionFilterEnabled(e.target.checked);
                if (!e.target.checked) setSelectedRegions([]);
              }}
              style={{ accentColor: "#60d0ff" }}
            />
            <span style={{ color: "rgba(100,150,200,0.7)", fontSize: 12 }}>On</span>
          </label>
        </div>

        {/* Expanded multi-select checkbox panel — only when enabled */}
        {regionFilterEnabled && (
          <div style={{ paddingTop: 14 }}>
            {availableRegions.length === 0 ? (
              <p style={{ color: "rgba(100,150,200,0.5)", fontSize: 12, margin: 0 }}>
                No region or state data found in the uploaded dataset. Ensure your file
                includes a <code style={{ color: "#60d0ff", margin: "0 3px" }}>region</code>
                or <code style={{ color: "#60d0ff" }}>state</code> column.
              </p>
            ) : (
              <>
                {/* Count + Select all / Clear controls */}
                <div style={{
                  display: "flex",
                  justifyContent: "space-between",
                  alignItems: "center",
                  marginBottom: 12,
                }}>
                  <span style={{ color: "rgba(100,150,200,0.55)", fontSize: 12 }}>
                    {selectedRegions.length} of {availableRegions.length} selected
                  </span>
                  <div style={{ display: "flex", gap: 8 }}>
                    {[
                      ["Select all", () => setSelectedRegions(availableRegions.map(e => e.value))],
                      ["Clear", () => setSelectedRegions([])],
                    ].map(([label, action]) => (
                      <button
                        key={label}
                        onClick={action}
                        style={{
                          background: "transparent",
                          border: "1px solid rgba(0,150,255,0.25)",
                          color: "#60d0ff",
                          borderRadius: 6,
                          padding: "3px 10px",
                          fontSize: 11,
                          cursor: "pointer",
                        }}
                      >
                        {label}
                      </button>
                    ))}
                  </div>
                </div>

                {/* Checkbox grid — auto-fills columns based on available width */}
                <div style={{
                  display: "grid",
                  gridTemplateColumns: "repeat(auto-fill, minmax(180px, 1fr))",
                  gap: 8,
                }}>
                  {availableRegions.map(entry => {
                    const checked = selectedRegions.includes(entry.value);
                    return (
                      <label
                        key={entry.type + "::" + entry.value}
                        style={{
                          display: "flex",
                          alignItems: "center",
                          gap: 8,
                          padding: "8px 12px",
                          background: checked ? "rgba(0,150,255,0.12)" : "rgba(0,30,80,0.5)",
                          border: `1px solid ${checked ? "rgba(0,150,255,0.45)" : "rgba(0,150,255,0.12)"}`,
                          borderRadius: 8,
                          cursor: "pointer",
                          transition: "all 0.15s",
                        }}
                      >
                        <input
                          type="checkbox"
                          checked={checked}
                          onChange={() => toggleRegion(entry.value)}
                          style={{ accentColor: "#60d0ff", flexShrink: 0 }}
                        />
                        <div style={{ overflow: "hidden", minWidth: 0 }}>
                          <span style={{
                            color: checked ? "#e2e8f0" : "rgba(100,150,200,0.7)",
                            fontSize: 13,
                            fontWeight: checked ? 600 : 400,
                            display: "block",
                            overflow: "hidden",
                            textOverflow: "ellipsis",
                            whiteSpace: "nowrap",
                          }}>
                            {entry.label}
                          </span>
                          <span style={{
                            fontSize: 10,
                            color: entry.type === "state" ? "rgba(100,200,180,0.6)" : "rgba(180,140,255,0.6)",
                            fontFamily: "'DM Mono', monospace",
                            letterSpacing: 0.5,
                            textTransform: "uppercase",
                          }}>
                            {entry.type}
                          </span>
                        </div>
                      </label>
                    );
                  })}
                </div>
              </>
            )}
          </div>
        )}
      </Card>

      {/* ── Existing numeric constraints ── */}
      <Card style={{ marginBottom: 24 }} glow>
        {constraints.map((c, i) => (
          <div key={c.key} style={{
            display: "grid",
            gridTemplateColumns: "1fr 120px 100px 80px",
            gap: 12,
            alignItems: "center",
            padding: "14px 0",
            borderBottom: i < constraints.length - 1 ? "1px solid rgba(0,150,255,0.1)" : "none",
            opacity: c.enabled ? 1 : 0.45,
            transition: "opacity 0.2s",
          }}>
            <span style={{ color: "#e2e8f0", fontSize: 14, fontWeight: 600 }}>{c.label}</span>
            <select
              value={c.operator}
              onChange={(e) => onChange(i, "operator", e.target.value)}
              style={sharedInputStyle}
            >
              <option value="gte">≥</option>
              <option value="lte">≤</option>
              <option value="eq">=</option>
            </select>
            <input
              type="number"
              value={c.value}
              onChange={(e) => onChange(i, "value", parseFloat(e.target.value) || 0)}
              style={{ ...sharedInputStyle, width: "100%" }}
            />
            <label style={{ display: "flex", alignItems: "center", gap: 6, cursor: "pointer" }}>
              <input
                type="checkbox"
                checked={c.enabled}
                onChange={(e) => onChange(i, "enabled", e.target.checked)}
                style={{ accentColor: "#60d0ff" }}
              />
              <span style={{ color: "rgba(100,150,200,0.7)", fontSize: 12 }}>On</span>
            </label>
          </div>
        ))}
      </Card>

      <GlowButton onClick={handleNext}>Run Optimization →</GlowButton>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// AHP MATRIX EDITOR
// ─────────────────────────────────────────────────────────────────────────────

function AHPMatrix({ matrix, onChange }) {
  const [open, setOpen] = useState(false);
  const labels = CRITERIA.map(c => c.label.split(" ")[0]);
  return (
    <Card style={{ marginBottom: 24 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <SectionTitle>AHP Pairwise Matrix</SectionTitle>
        <button onClick={() => setOpen(!open)} style={{
          background: "rgba(0,100,200,0.15)",
          border: "1px solid rgba(0,150,255,0.3)",
          color: "#60d0ff",
          borderRadius: 6, padding: "4px 14px", fontSize: 11, cursor: "pointer",
          fontFamily: "'DM Mono', monospace", letterSpacing: 1
        }}>{open ? "HIDE" : "EDIT"}</button>
      </div>
      {open && (
        <div style={{ overflowX: "auto" }}>
          <table style={{ fontSize: 12, borderCollapse: "collapse" }}>
            <thead>
              <tr>
                <th style={{ padding: "4px 8px", color: "rgba(100,180,255,0.5)" }}></th>
                {labels.map(l => <th key={l} style={{ padding: "4px 8px", color: "rgba(100,180,255,0.6)", fontFamily: "'DM Mono', monospace", fontSize: 10 }}>{l}</th>)}
              </tr>
            </thead>
            <tbody>
              {matrix.map((row, i) => (
                <tr key={i}>
                  <td style={{ padding: "4px 8px", color: "rgba(100,180,255,0.6)", fontWeight: 600, fontFamily: "'DM Mono', monospace", fontSize: 10 }}>{labels[i]}</td>
                  {row.map((val, j) => (
                    <td key={j} style={{ padding: 2 }}>
                      <input
                        type="number" step="0.1" value={val}
                        onChange={(e) => {
                          const nv = parseFloat(e.target.value) || 0;
                          const nm = matrix.map((r, ri) => r.map((v, ci) => {
                            if (ri === i && ci === j) return nv;
                            if (ri === j && ci === i) return nv !== 0 ? 1 / nv : 0;
                            return v;
                          }));
                          onChange(nm);
                        }}
                        style={{
                          width: 52,
                          background: i === j ? "rgba(0,10,30,0.8)" : "rgba(0,30,80,0.6)",
                          border: "1px solid rgba(0,150,255,0.2)",
                          color: "#60d0ff",
                          borderRadius: 4, padding: "4px 6px", fontSize: 12, textAlign: "center",
                          outline: "none"
                        }}
                        disabled={i === j}
                      />
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </Card>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// STEP 4 — OPTIMIZATION
// ─────────────────────────────────────────────────────────────────────────────

function Optimization({
  locations,
  constraints,
  // ── CHANGE 4: accept regionFilter from parent ──
  regionFilter = { regionFilterEnabled: false, selectedRegions: [] },
  onNext,
  onMatrixChange,
}) {
  const [matrix, setMatrix] = useState(DEFAULT_MATRIX);
  const [results, setResults] = useState(null);
  const [weights, setWeights] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);

  const handleMatrixChange = (newMatrix) => {
    setMatrix(newMatrix);
    if (onMatrixChange) onMatrixChange(newMatrix);
  };

  const run = async () => {
    setLoading(true); setError(null);
    try {
      // ── Apply region/state filter before sending to backend ──
      const { regionFilterEnabled, selectedRegions } = regionFilter;
      const filteredLocations =
        regionFilterEnabled && selectedRegions.length > 0
          ? locations.filter(l => {
            // A location passes if its region OR its state is in the selected list.
            // Checking both independently so state-only selections also work.
            const matchesRegion = l.region && selectedRegions.includes(l.region);
            const matchesState = l.state && selectedRegions.includes(l.state);
            return matchesRegion || matchesState;
          })
          : locations;

      const r = await fetch(`${API}/api/analyze`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          locations: filteredLocations,
          pairwiseMatrix: matrix,
          constraints,
        }),
      });
      const d = await r.json();
      setWeights(d.weights);
      setResults(d.results);
    } catch (e) {
      setError("Backend error: " + e.message);
    } finally { setLoading(false); }
  };

  const handleNext = () => onNext(results, weights, matrix);

  const feasible = results?.filter(r => r.feasible) ?? [];
  const infeasible = results?.filter(r => !r.feasible) ?? [];

  return (
    <div style={{ maxWidth: 900, margin: "0 auto" }}>
      <SectionTitle>MCDM Optimization — AHP + Entropy + TOPSIS</SectionTitle>

      {/* Show active region filter summary if set */}
      {regionFilter.regionFilterEnabled && regionFilter.selectedRegions.length > 0 && (
        <div style={{
          background: "rgba(0,80,160,0.15)",
          border: "1px solid rgba(0,150,255,0.25)",
          borderRadius: 10,
          padding: "10px 16px",
          marginBottom: 16,
          fontSize: 12,
          color: "rgba(100,180,255,0.8)",
          fontFamily: "'DM Mono', monospace",
          display: "flex",
          alignItems: "center",
          gap: 10,
          flexWrap: "wrap",
        }}>
          <span style={{ color: "rgba(100,150,200,0.6)" }}>REGION FILTER ACTIVE:</span>
          {regionFilter.selectedRegions.map(r => (
            <span key={r} style={{
              background: "rgba(0,150,255,0.15)",
              border: "1px solid rgba(0,150,255,0.3)",
              borderRadius: 4,
              padding: "2px 8px",
              color: "#60d0ff",
              display: "inline-flex",
              alignItems: "center",
              gap: 5,
            }}>
              {r}
            </span>
          ))}
        </div>
      )}

      <AHPMatrix matrix={matrix} onChange={handleMatrixChange} />

      {!results && !loading && (
        <GlowButton onClick={run} style={{ marginBottom: 24 }}>Run Analysis</GlowButton>
      )}

      {loading && <Spinner />}
      {error && (
        <div style={{ color: "#ff6b6b", padding: 12, fontFamily: "'DM Mono', monospace", fontSize: 12 }}>
          {error}
        </div>
      )}

      {results && !loading && (
        <>
          {weights && (
            <Card style={{ marginBottom: 24 }} glow>
              <SectionTitle>Hybrid Weights (AHP + Entropy)</SectionTitle>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 12 }}>
                {weights.map(w => (
                  <div key={w.key} style={{
                    background: "rgba(0,20,60,0.6)", borderRadius: 10, padding: 14,
                    border: "1px solid rgba(0,150,255,0.15)"
                  }}>
                    <p style={{ color: "rgba(100,150,200,0.6)", fontSize: 10, margin: "0 0 6px", letterSpacing: 1.5, fontFamily: "'DM Mono', monospace" }}>
                      {w.name} {w.isCost ? "🔻" : "🔺"}
                    </p>
                    <div style={{ display: "flex", gap: 8, alignItems: "baseline" }}>
                      <span style={{ color: "#60d0ff", fontSize: 20, fontWeight: 800 }}>
                        {pct(w.combinedWeight)}
                      </span>
                    </div>
                    <p style={{ color: "rgba(60,100,150,0.7)", fontSize: 10, margin: "4px 0 0", fontFamily: "'DM Mono', monospace" }}>
                      AHP {pct(w.ahpWeight)} · Entropy {pct(w.entropyWeight)}
                    </p>
                  </div>
                ))}
              </div>
            </Card>
          )}

          <Card style={{ marginBottom: 24 }}>
            <SectionTitle>Location Rankings</SectionTitle>
            <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
              {feasible.map(r => (
                <div key={r.locationId} style={{
                  background: "rgba(0,20,60,0.5)", borderRadius: 10, padding: "14px 18px",
                  border: r.rank === 1 ? "1px solid rgba(0,200,255,0.4)" : "1px solid rgba(0,100,200,0.15)",
                  display: "grid", gridTemplateColumns: "40px 1fr 200px 80px", gap: 16, alignItems: "center",
                  boxShadow: r.rank === 1 ? "0 0 20px rgba(0,200,255,0.1)" : "none"
                }}>
                  <Badge rank={r.rank} />
                  <div>
                    <p style={{ margin: 0, fontWeight: 700, color: "#e2e8f0", fontSize: 15 }}>{r.locationName}</p>
                    <p style={{ margin: "2px 0 0", color: "rgba(100,150,200,0.5)", fontSize: 10, fontFamily: "'DM Mono', monospace" }}>
                      {Object.entries(r.criteriaScores).map(([k, v]) =>
                        `${CRITERIA.find(c => c.key === k)?.label?.split(" ")[0]}: ${fmt(v, 3)}`
                      ).join(" · ")}
                    </p>
                  </div>
                  <ScoreBar score={r.compositeScore} />
                  <Chip
                    label={r.rank === 1 ? "⭐ Optimal" : `Rank #${r.rank}`}
                    color={r.rank === 1 ? "rgba(0,80,80,0.8)" : "rgba(0,30,80,0.6)"}
                  />
                </div>
              ))}
              {infeasible.map(r => (
                <div key={r.locationId} style={{
                  background: "rgba(0,20,60,0.3)", borderRadius: 10, padding: "10px 18px",
                  border: "1px solid rgba(0,100,200,0.1)", opacity: 0.5,
                  display: "flex", alignItems: "center", gap: 12
                }}>
                  <span style={{ color: "#ff6b6b", fontSize: 16 }}>✗</span>
                  <span style={{ color: "rgba(100,150,200,0.6)", fontSize: 14 }}>{r.locationName}</span>
                  <Chip label="Infeasible" color="rgba(80,20,20,0.8)" />
                </div>
              ))}
            </div>
          </Card>

          <GlowButton onClick={handleNext}>Run Monte Carlo Simulation →</GlowButton>
        </>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// STEP 5 — SIMULATION
// ─────────────────────────────────────────────────────────────────────────────

function Simulation({ locations, weights, constraints, regionFilter = { regionFilterEnabled: false, selectedRegions: [] }, onNext }) {
  const [iter, setIter] = useState(1000);
  const [results, setResults] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);

  // ── Apply the same feasibility filter as Optimization ────────────────────
  // Must mirror Optimization.run() exactly so Monte Carlo and MCDM are consistent.
  const feasibleLocations = useMemo(() => {
    let locs = locations;

    // 1. Region / state filter
    const { regionFilterEnabled, selectedRegions } = regionFilter;
    if (regionFilterEnabled && selectedRegions.length > 0) {
      locs = locs.filter(l => {
        const matchesRegion = l.region && selectedRegions.includes(l.region);
        const matchesState = l.state && selectedRegions.includes(l.state);
        return matchesRegion || matchesState;
      });
    }

    // 2. Numeric constraints
    const active = constraints.filter(c => c.enabled);
    if (active.length > 0) {
      locs = locs.filter(l =>
        active.every(c => {
          const val = parseFloat(l[c.key]) || 0;
          if (c.operator === "gte") return val >= c.value;
          if (c.operator === "lte") return val <= c.value;
          if (c.operator === "eq") return Math.abs(val - c.value) <= 0.001;
          return true;
        })
      );
    }
    return locs;
  }, [locations, constraints, regionFilter]);

  const activeConstraints = constraints.filter(c => c.enabled);
  const { regionFilterEnabled, selectedRegions } = regionFilter;
  const hasFilter = (regionFilterEnabled && selectedRegions.length > 0) || activeConstraints.length > 0;
  const filteredOut = locations.length - feasibleLocations.length;

  const run = async () => {
    if (feasibleLocations.length === 0) {
      setError("No feasible locations match the current constraints. Please adjust your constraints in Step 3.");
      return;
    }
    setLoading(true); setError(null);
    try {
      const r = await fetch(`${API}/api/monte-carlo`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          // Send only the pre-filtered feasible set so the backend gets exact same input
          locations: feasibleLocations,
          weights,
          iterations: iter,
          constraints,          // backend still validates numeric constraints as defence
          regionFilter,         // backend also applies region filter as defence-in-depth
        }),
      });
      const d = await r.json();
      setResults(d.monteCarloResults);
    } catch (e) {
      setError("Backend error: " + e.message);
    } finally { setLoading(false); }
  };

  return (
    <div style={{ maxWidth: 860, margin: "0 auto" }}>
      <SectionTitle>Monte Carlo Robustness Simulation</SectionTitle>

      {/* ── Active filter badge ── */}
      {hasFilter && (
        <div style={{
          background: "rgba(0,80,160,0.15)",
          border: "1px solid rgba(0,150,255,0.25)",
          borderRadius: 10,
          padding: "10px 16px",
          marginBottom: 16,
          fontSize: 12,
          color: "rgba(100,180,255,0.8)",
          fontFamily: "'DM Mono', monospace",
          display: "flex",
          alignItems: "center",
          gap: 12,
          flexWrap: "wrap",
        }}>
          <span style={{ color: "rgba(100,150,200,0.6)" }}>CONSTRAINTS ACTIVE:</span>
          {regionFilterEnabled && selectedRegions.length > 0 && (
            <span style={{ color: "#60d0ff" }}>Region/State filter ({selectedRegions.length} selected)</span>
          )}
          {activeConstraints.map(c => (
            <span key={c.key} style={{
              background: "rgba(0,150,255,0.15)", border: "1px solid rgba(0,150,255,0.3)",
              borderRadius: 4, padding: "2px 8px", color: "#60d0ff",
            }}>
              {c.label} {c.operator === "gte" ? "≥" : c.operator === "lte" ? "≤" : "="} {c.value}
            </span>
          ))}
          <span style={{ marginLeft: "auto", color: filteredOut > 0 ? "#f5a623" : "#4ade80", fontWeight: 700 }}>
            {feasibleLocations.length}/{locations.length} feasible
          </span>
        </div>
      )}

      {/* ── Zero feasible locations warning ── */}
      {feasibleLocations.length === 0 && (
        <div style={{
          background: "rgba(255,80,80,0.08)", border: "1px solid rgba(255,80,80,0.25)",
          borderRadius: 12, padding: "20px 24px", marginBottom: 24,
          color: "#ff6b6b", fontFamily: "'DM Mono', monospace", fontSize: 13,
        }}>
          ⚠ No locations pass the current constraints — adjust them in Step 3 before running simulation.
        </div>
      )}

      <Card style={{ marginBottom: 24 }} glow>
        <p style={{ color: "rgba(150,190,255,0.7)", fontSize: 13, margin: "0 0 16px" }}>
          Perturbs weights (±15%) and criteria values (±5%) over{" "}
          <strong style={{ color: "#60d0ff" }}>{iter.toLocaleString()}</strong> iterations across{" "}
          <strong style={{ color: feasibleLocations.length > 0 ? "#4ade80" : "#f5a623" }}>
            {feasibleLocations.length} feasible location{feasibleLocations.length !== 1 ? "s" : ""}
          </strong>{" "}to assess ranking stability.
        </p>
        <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
          <span style={{ color: "rgba(100,150,200,0.7)", fontSize: 12, fontFamily: "'DM Mono', monospace" }}>ITERATIONS</span>
          {[500, 1000, 2000, 5000].map(n => (
            <button key={n} onClick={() => setIter(n)} style={{
              background: iter === n ? "rgba(0,150,255,0.2)" : "rgba(0,30,80,0.6)",
              color: iter === n ? "#60d0ff" : "rgba(100,150,200,0.6)",
              border: iter === n ? "1px solid rgba(0,180,255,0.5)" : "1px solid rgba(0,100,200,0.2)",
              borderRadius: 8, padding: "7px 16px", fontSize: 12,
              cursor: "pointer", fontFamily: "'DM Mono', monospace",
              boxShadow: iter === n ? "0 0 15px rgba(0,180,255,0.2)" : "none",
              transition: "all 0.2s"
            }}>{n.toLocaleString()}</button>
          ))}
          <GlowButton
            onClick={run}
            style={{ marginLeft: "auto", opacity: feasibleLocations.length === 0 ? 0.4 : 1 }}
          >Run Simulation</GlowButton>
        </div>
      </Card>

      {loading && <Spinner />}
      {error && <div style={{ color: "#ff6b6b", padding: 12, fontFamily: "'DM Mono', monospace", fontSize: 12 }}>{error}</div>}

      {results && !loading && (
        <>
          <Card style={{ marginBottom: 24 }} glow>
            <SectionTitle>Rank Stability Analysis</SectionTitle>
            {results.length === 0 ? (
              <p style={{ color: "rgba(100,150,200,0.6)", fontSize: 13, fontFamily: "'DM Mono', monospace" }}>
                No results — all locations were filtered out by constraints.
              </p>
            ) : (
              results.map((r) => (
                <div key={r.locationId} style={{
                  background: "rgba(0,20,60,0.5)", borderRadius: 10, padding: 16, marginBottom: 10,
                  border: "1px solid rgba(0,100,200,0.15)"
                }}>
                  <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 10 }}>
                    <div>
                      <span style={{ fontWeight: 700, color: "#e2e8f0" }}>{r.locationName}</span>
                      <span style={{ color: "rgba(100,150,200,0.5)", fontSize: 12, marginLeft: 10, fontFamily: "'DM Mono', monospace" }}>
                        Avg Rank: <strong style={{ color: "#f5a623" }}>{r.avgRank.toFixed(2)}</strong>
                      </span>
                      <span style={{ color: "rgba(100,150,200,0.5)", fontSize: 12, marginLeft: 10, fontFamily: "'DM Mono', monospace" }}>
                        CI: [{r.confidenceInterval[0]}–{r.confidenceInterval[1]}]
                      </span>
                    </div>
                    <Badge rank={Math.round(r.avgRank)} />
                  </div>
                  <div style={{ display: "flex", gap: 4 }}>
                    {r.rankProbabilities.map((p, i) => (
                      <div key={i} style={{ flex: 1, textAlign: "center" }}>
                        <div style={{
                          height: Math.max(4, p * 80),
                          background: p > 0.5 ? "rgba(0,180,255,0.7)" : p > 0.2 ? "#f5a623" : "rgba(0,60,120,0.6)",
                          borderRadius: "3px 3px 0 0", transition: "height .4s",
                          boxShadow: p > 0.5 ? "0 0 8px rgba(0,180,255,0.4)" : "none"
                        }} />
                        <div style={{ color: "rgba(100,150,200,0.4)", fontSize: 9, marginTop: 2, fontFamily: "'DM Mono', monospace" }}>#{i + 1}</div>
                      </div>
                    ))}
                  </div>
                </div>
              ))
            )}
          </Card>
          <GlowButton onClick={onNext}>Proceed to Export →</GlowButton>
        </>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// STEP 6 — EXPORT
// ─────────────────────────────────────────────────────────────────────────────

function Export({ analysisResults, weights, locations, constraints, pairwiseMatrix, regionFilter }) {
  const [exporting, setExporting] = useState(false);
  const [exportError, setExportError] = useState(null);

  const exportCSV = () => {
    if (!analysisResults) return;
    const rows = [
      ["Rank", "Location", "Composite Score", "Feasible", ...CRITERIA.map(c => c.label)],
      ...analysisResults.map(r => [
        r.rank, r.locationName, r.compositeScore, r.feasible ? "Yes" : "No",
        ...CRITERIA.map(c => r.criteriaScores[c.key] ?? "N/A")
      ])
    ];
    const csv = rows.map(r => r.join(",")).join("\n");
    const blob = new Blob([csv], { type: "text/csv" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "Ashok_Leyland_Plant_Location_Ranking.csv";
    a.click();
    URL.revokeObjectURL(a.href);
  };

  const exportJSON = () => {
    const data = {
      analysisResults, weights, locations, constraints, pairwiseMatrix,
      exportedAt: new Date().toISOString()
    };
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "Ashok_Leyland_Plant_Location_Results.json";
    a.click();
    URL.revokeObjectURL(a.href);
  };

  const exportFullExcel = async () => {
    if (!analysisResults || !locations || !weights) {
      alert("No analysis results available yet. Please run the Optimization step first.");
      return;
    }
    setExporting(true); setExportError(null);
    try {
      const response = await fetch(`${API}/api/export-excel`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          locations, results: analysisResults, weights,
          pairwiseMatrix: pairwiseMatrix ?? [],
          constraints: constraints ?? [],
          regionFilter: regionFilter ?? { regionFilterEnabled: false, selectedRegions: [] },
        }),
      });
      if (!response.ok) {
        let msg = `Server error ${response.status}`;
        try { const j = await response.json(); msg = j.error ?? msg; } catch (_) { }
        throw new Error(msg);
      }
      const arrayBuffer = await response.arrayBuffer();
      const magic = new Uint8Array(arrayBuffer, 0, 4);
      if (!(magic[0] === 0x50 && magic[1] === 0x4B)) {
        let detail = "Backend returned unexpected content instead of an xlsx file.";
        try { detail = JSON.parse(new TextDecoder().decode(arrayBuffer))?.error ?? detail; } catch (_) { }
        throw new Error(detail);
      }
      const dateStr = new Date().toISOString().slice(0, 10);
      const blob = new Blob([arrayBuffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `Ashok_Leyland_Plant_Location_Results_${dateStr}.xlsx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (err) {
      console.error("Excel export error:", err);
      setExportError("Export failed: " + err.message);
    } finally { setExporting(false); }
  };

  const top = analysisResults?.find(r => r.rank === 1);

  return (
    <div style={{ maxWidth: 1000, margin: "0 auto" }}>
      <SectionTitle>Export Results</SectionTitle>

      {exportError && (
        <div style={{
          background: "rgba(255,50,50,0.1)", border: "1px solid rgba(255,50,50,0.3)",
          borderRadius: 10, padding: 12, marginBottom: 20,
          color: "#ff6b6b", fontSize: 13, fontFamily: "'DM Mono', monospace"
        }}>❌ {exportError}</div>
      )}

      {top && (
        <div style={{
          background: "linear-gradient(135deg, rgba(0,20,60,0.9), rgba(0,40,80,0.7))",
          border: "1px solid rgba(0,200,255,0.35)", borderRadius: 16, padding: 24, marginBottom: 24,
          boxShadow: "0 0 40px rgba(0,180,255,0.1)"
        }}>
          <p style={{
            color: "rgba(0,200,255,0.7)", fontSize: 10, letterSpacing: 3, fontWeight: 700,
            margin: "0 0 8px", fontFamily: "'DM Mono', monospace"
          }}>⭐ OPTIMAL PLANT LOCATION</p>
          <h2 style={{
            color: "#fff", fontSize: 28, margin: "0 0 4px",
            fontFamily: "Georgia, serif", fontWeight: 400,
            textShadow: "0 0 30px rgba(100,200,255,0.3)"
          }}>{top.locationName}</h2>
          <p style={{ color: "rgba(150,200,255,0.6)", margin: 0, fontSize: 13, fontFamily: "'DM Mono', monospace" }}>
            TOPSIS Score: <strong style={{ color: "#60d0ff" }}>{fmt(top.compositeScore)}</strong>
          </p>
        </div>
      )}

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 16, marginBottom: 24 }}>
        <button onClick={exportFullExcel} disabled={exporting} style={{
          background: exporting ? "rgba(0,80,40,0.3)" : "linear-gradient(135deg, rgba(0,100,0,0.3), rgba(0,150,0,0.2))",
          border: exporting ? "1px solid rgba(0,200,100,0.2)" : "1px solid rgba(0,200,100,0.4)",
          color: exporting ? "#90ee90" : "#98fb98",
          borderRadius: 14, padding: 20, cursor: exporting ? "wait" : "pointer",
          textAlign: "left", transition: "all 0.25s", opacity: exporting ? 0.7 : 1,
          position: "relative", overflow: "hidden"
        }}
          onMouseEnter={e => { if (!exporting) { e.currentTarget.style.background = "linear-gradient(135deg, rgba(0,150,0,0.4), rgba(0,200,0,0.3))"; e.currentTarget.style.borderColor = "rgba(0,255,100,0.6)"; e.currentTarget.style.boxShadow = "0 0 25px rgba(0,255,100,0.2)"; } }}
          onMouseLeave={e => { if (!exporting) { e.currentTarget.style.background = "linear-gradient(135deg, rgba(0,100,0,0.3), rgba(0,150,0,0.2))"; e.currentTarget.style.borderColor = "rgba(0,200,100,0.4)"; e.currentTarget.style.boxShadow = "none"; } }}
        >
          {exporting && (
            <div style={{
              position: "absolute", top: 0, left: 0, right: 0, bottom: 0,
              background: "linear-gradient(90deg, transparent, rgba(255,255,255,0.1), transparent)",
              animation: "shimmer 1.5s infinite"
            }} />
          )}
          <div style={{ fontSize: 24, marginBottom: 8 }}>{exporting ? "⏳" : "📊"}</div>
          <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 4 }}>{exporting ? "GENERATING..." : "Export Complete Excel"}</div>
          <div style={{ color: "rgba(144,238,144,0.6)", fontSize: 12 }}>{exporting ? "Creating 6‑sheet workbook" : "6 sheets · Constraints · Raw data · Weights · Rankings · Summary · Dashboard"}</div>
        </button>

        <button onClick={exportCSV} style={{
          background: "rgba(0,30,80,0.5)", border: "1px solid rgba(0,150,255,0.25)",
          color: "#60d0ff", borderRadius: 14, padding: 20, cursor: "pointer", textAlign: "left", transition: "all 0.25s",
        }}
          onMouseEnter={e => { e.currentTarget.style.background = "rgba(0,60,120,0.5)"; e.currentTarget.style.borderColor = "rgba(0,200,255,0.4)"; e.currentTarget.style.boxShadow = "0 0 20px rgba(0,180,255,0.15)"; }}
          onMouseLeave={e => { e.currentTarget.style.background = "rgba(0,30,80,0.5)"; e.currentTarget.style.borderColor = "rgba(0,150,255,0.25)"; e.currentTarget.style.boxShadow = "none"; }}
        >
          <div style={{ fontSize: 24, marginBottom: 8 }}>📈</div>
          <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 4 }}>Export CSV</div>
          <div style={{ color: "rgba(100,160,220,0.6)", fontSize: 12 }}>Rankings & scores only</div>
        </button>

        <button onClick={exportJSON} style={{
          background: "rgba(80,20,80,0.5)", border: "1px solid rgba(200,100,255,0.25)",
          color: "#d896ff", borderRadius: 14, padding: 20, cursor: "pointer", textAlign: "left", transition: "all 0.25s",
        }}
          onMouseEnter={e => { e.currentTarget.style.background = "rgba(120,20,120,0.5)"; e.currentTarget.style.borderColor = "rgba(255,100,255,0.4)"; e.currentTarget.style.boxShadow = "0 0 20px rgba(255,100,255,0.15)"; }}
          onMouseLeave={e => { e.currentTarget.style.background = "rgba(80,20,80,0.5)"; e.currentTarget.style.borderColor = "rgba(200,100,255,0.25)"; e.currentTarget.style.boxShadow = "none"; }}
        >
          <div style={{ fontSize: 24, marginBottom: 8 }}>🗂️</div>
          <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 4 }}>Export JSON</div>
          <div style={{ color: "rgba(200,160,220,0.6)", fontSize: 12 }}>Full analysis payload</div>
        </button>
      </div>

      <Card glow style={{ marginBottom: 24 }}>
        <SectionTitle>Complete Excel Workbook Structure</SectionTitle>
        <div style={{ display: "grid", gap: 12 }}>
          {[
            { sheet: "0_Constraints", desc: "Active constraints, feasibility summary, and details of locations filtered out", icon: "🔒" },
            { sheet: "1_Full_Input_Data", desc: "Restores ALL parent + child sub-attributes from original upload (grouped by 7 sections)", icon: "📋" },
            { sheet: "2_Normalised_Matrix", desc: "Min-Max normalized decision matrix with all final scores (0 = worst, 1 = best)", icon: "📐" },
            { sheet: "3_Weight_Calculation", desc: "AHP pairwise matrix, Entropy weights, and Hybrid weights (60% AHP + 40% Entropy)", icon: "⚖️" },
            { sheet: "4_TOPSIS_Ranking", desc: "Weighted matrix, closeness scores, and final rankings", icon: "🏆" },
            { sheet: "5_Dashboard", desc: "Executive summary, location rankings table, and interactive bar charts", icon: "📊" },
          ].map((item, i) => (
            <div key={i} style={{
              display: "flex", alignItems: "center", gap: 12, padding: "10px 14px",
              background: "rgba(0,20,60,0.4)", borderRadius: 10, border: "1px solid rgba(0,150,255,0.1)"
            }}>
              <span style={{ fontSize: 20 }}>{item.icon}</span>
              <div style={{ flex: 1 }}>
                <span style={{ color: "#60d0ff", fontWeight: 700, fontSize: 13, fontFamily: "'DM Mono', monospace", marginRight: 12 }}>{item.sheet}</span>
                <span style={{ color: "rgba(150,200,255,0.7)", fontSize: 12 }}>{item.desc}</span>
              </div>
            </div>
          ))}
        </div>
      </Card>

      {analysisResults && (
        <Card glow>
          <SectionTitle>Final Rankings Summary</SectionTitle>
          {analysisResults.filter(r => r.feasible).slice(0, 5).map(r => (
            <div key={r.locationId} style={{
              display: "flex", alignItems: "center", gap: 14, padding: "10px 0",
              borderBottom: "1px solid rgba(0,100,200,0.12)"
            }}>
              <Badge rank={r.rank} />
              <span style={{ flex: 1, color: "#e2e8f0", fontWeight: 600 }}>{r.locationName}</span>
              <ScoreBar score={r.compositeScore} />
            </div>
          ))}
        </Card>
      )}

      <style>{`
        @keyframes shimmer {
          0% { transform: translateX(-100%); }
          100% { transform: translateX(100%); }
        }
      `}</style>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// MAIN APP
// ─────────────────────────────────────────────────────────────────────────────

export default function App() {
  const [step, setStep] = useState(0);
  const [locations, setLocations] = useState([]);
  const [constraints, setConstraints] = useState(DEFAULT_CONSTRAINTS);
  const [analysisResults, setAR] = useState(null);
  const [weights, setWeights] = useState(null);
  const [pairwiseMatrix, setPairwiseMatrix] = useState(DEFAULT_MATRIX);
  // Lifted from DataInput so the preview survives back-navigation
  const [uploadPreview, setUploadPreview] = useState(null);
  // ── store region filter state set by Constraints step ──
  const [regionFilter, setRegionFilter] = useState({
    regionFilterEnabled: false,
    selectedRegions: [],
  });

  const updateConstraint = (i, field, val) => {
    setConstraints(prev => prev.map((c, idx) => idx === i ? { ...c, [field]: val } : c));
  };

  return (
    <div style={{
      minHeight: "100vh", background: "#020b1a", color: "#e2e8f0",
      fontFamily: "'DM Sans', 'Segoe UI', sans-serif", position: "relative", overflow: "hidden",
    }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700;800&family=DM+Mono:wght@400;500&display=swap');
        * { box-sizing: border-box; }
        @keyframes spin { to { transform: rotate(360deg); } }
        @keyframes pulse-glow { 0%, 100% { opacity: 0.4; } 50% { opacity: 0.7; } }
        @keyframes scanline { 0% { transform: translateY(-100%); } 100% { transform: translateY(100vh); } }
        input[type=number]::-webkit-inner-spin-button { opacity: 1; }
        ::-webkit-scrollbar { width: 5px; height: 5px; }
        ::-webkit-scrollbar-track { background: rgba(0,20,60,0.3); }
        ::-webkit-scrollbar-thumb { background: rgba(0,150,255,0.2); border-radius: 3px; }
        select option { background: #021030; }
      `}</style>

      <div style={{
        position: "fixed", inset: 0, pointerEvents: "none", zIndex: 0,
        background: `
          radial-gradient(ellipse 80% 60% at 50% -20%, rgba(0,80,200,0.25) 0%, transparent 70%),
          radial-gradient(ellipse 60% 50% at 80% 80%, rgba(0,40,120,0.15) 0%, transparent 60%),
          radial-gradient(ellipse 40% 40% at 20% 60%, rgba(0,60,160,0.1) 0%, transparent 60%)
        `
      }} />
      <div style={{
        position: "fixed", inset: 0, pointerEvents: "none", zIndex: 0,
        backgroundImage: `
          linear-gradient(rgba(0,100,255,0.03) 1px, transparent 1px),
          linear-gradient(90deg, rgba(0,100,255,0.03) 1px, transparent 1px)
        `,
        backgroundSize: "40px 40px"
      }} />

      {/* Background Logo Watermark */}
      <div style={{
        position: "fixed", inset: 0, pointerEvents: "none", zIndex: 0,
        backgroundImage: `url(${logo})`,
        backgroundRepeat: "no-repeat",
        backgroundPosition: "center center",
        backgroundSize: "1090px auto",
        opacity: 0.22,
        filter: "blur(0px)",
      }} />

      <header style={{
        borderBottom: "1px solid rgba(0,150,255,0.12)", padding: "0 32px",
        display: "flex", alignItems: "center", gap: 20, height: 64,
        position: "sticky", top: 0, zIndex: 100,
        background: "rgba(2,10,28,0.92)", backdropFilter: "blur(16px)",
        boxShadow: "0 1px 30px rgba(0,100,255,0.08)"
      }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <div style={{
            width: 36, height: 36,
            background: "linear-gradient(135deg, rgba(0,100,255,0.3), rgba(0,200,255,0.2))",
            border: "1px solid rgba(0,180,255,0.4)", borderRadius: 8,
            display: "flex", alignItems: "center", justifyContent: "center",
            fontWeight: 900, fontSize: 13, color: "#60d0ff",
            fontFamily: "'DM Mono', monospace", boxShadow: "0 0 16px rgba(0,160,255,0.3)"
          }}>AL</div>
          <div>
            <div style={{ fontWeight: 800, fontSize: 15, letterSpacing: 0.5, color: "#e2e8f0" }}>Ashok Leyland</div>
            <div style={{ fontSize: 9, color: "rgba(0,180,255,0.5)", letterSpacing: 2.5, textTransform: "uppercase", fontFamily: "'DM Mono', monospace" }}>
              Plant Location Decision System
            </div>
          </div>
        </div>
        <div style={{ flex: 1 }} />
        <div style={{ fontSize: 10, color: "rgba(0,150,255,0.4)", fontFamily: "'DM Mono', monospace", letterSpacing: 1.5 }}>
          Operations Research · AHP · TOPSIS · Monte Carlo
        </div>
      </header>

      <div style={{
        background: "rgba(2,8,24,0.85)", borderBottom: "1px solid rgba(0,120,255,0.1)",
        padding: "0 32px", backdropFilter: "blur(8px)", overflowX: "auto",
      }}>
        <div style={{ display: "flex", alignItems: "stretch", minWidth: "fit-content" }}>
          {STEPS.map((s, i) => {
            const active = i === step;
            const complete = i < step;
            return (
              <button key={s.id} onClick={() => i < step && setStep(i)} style={{
                display: "flex", alignItems: "center", gap: 10, padding: "14px 20px",
                background: "transparent", border: "none",
                borderBottom: active ? "2px solid rgba(0,200,255,0.8)" : "2px solid transparent",
                cursor: i < step ? "pointer" : "default",
                opacity: i > step ? 0.3 : 1, transition: "all .2s",
                whiteSpace: "nowrap", position: "relative",
              }}>
                {active && (
                  <div style={{
                    position: "absolute", bottom: 0, left: "20%", right: "20%",
                    height: 2, background: "rgba(0,200,255,0.8)",
                    boxShadow: "0 0 10px rgba(0,200,255,0.8)",
                  }} />
                )}
                <div style={{
                  width: 28, height: 28, borderRadius: "50%",
                  background: active ? "linear-gradient(135deg, rgba(0,140,255,0.3), rgba(0,200,255,0.15))" : complete ? "rgba(0,80,40,0.5)" : "rgba(0,30,80,0.5)",
                  border: active ? "1px solid rgba(0,200,255,0.6)" : complete ? "1px solid rgba(0,200,100,0.4)" : "1px solid rgba(0,80,160,0.3)",
                  display: "flex", alignItems: "center", justifyContent: "center", fontSize: 12,
                  color: active ? "#60d0ff" : complete ? "#4ade80" : "rgba(100,150,200,0.4)",
                  fontWeight: 700, flexShrink: 0,
                  boxShadow: active ? "0 0 12px rgba(0,200,255,0.25)" : "none"
                }}>
                  {complete ? "✓" : s.icon}
                </div>
                <div style={{ textAlign: "left" }}>
                  <div style={{
                    fontSize: 12, fontWeight: active ? 700 : 500,
                    color: active ? "#e2e8f0" : complete ? "rgba(150,200,255,0.6)" : "rgba(100,130,180,0.5)"
                  }}>{s.label}</div>
                  <div style={{ fontSize: 9, color: "rgba(0,150,255,0.3)", fontFamily: "'DM Mono', monospace", letterSpacing: 0.5 }}>{s.sub}</div>
                </div>
                {i < STEPS.length - 1 && (
                  <span style={{ color: "rgba(0,100,200,0.3)", marginLeft: 12, fontSize: 14 }}>›</span>
                )}
              </button>
            );
          })}
        </div>
      </div>

      <main style={{ padding: "40px 32px", maxWidth: 1400, margin: "0 auto", position: "relative", zIndex: 1, transition: "max-width 0.4s ease" }}>
        {step === 0 && <DataInput savedPreview={uploadPreview} onPreviewChange={setUploadPreview} onNext={(locs) => { setLocations(locs); setStep(1); }} />}
        {step === 1 && <Processing locations={locations} onNext={() => setStep(2)} />}

        {/* ── CHANGE 2: pass locations + capture regionFilter from onNext ── */}
        {step === 2 && (
          <Constraints
            constraints={constraints}
            onChange={updateConstraint}
            locations={locations}
            onNext={(rf) => { setRegionFilter(rf); setStep(3); }}
          />
        )}

        {/* ── CHANGE 3: pass regionFilter down to Optimization ── */}
        {step === 3 && (
          <Optimization
            locations={locations}
            constraints={constraints}
            regionFilter={regionFilter}
            onMatrixChange={setPairwiseMatrix}
            onNext={(results, w, matrix) => {
              setAR(results);
              setWeights(w);
              setPairwiseMatrix(matrix || pairwiseMatrix);
              setStep(4);
            }}
          />
        )}

        {step === 4 && (
          <Simulation
            locations={locations}
            weights={weights}
            constraints={constraints}
            regionFilter={regionFilter}
            onNext={() => setStep(5)}
          />
        )}
        {step === 5 && (
          <Export
            analysisResults={analysisResults}
            weights={weights}
            locations={locations}
            constraints={constraints}
            pairwiseMatrix={pairwiseMatrix}
            regionFilter={regionFilter}
          />
        )}
      </main>

      <footer style={{
        borderTop: "1px solid rgba(0,100,200,0.1)", padding: "14px 32px",
        display: "flex", justifyContent: "space-between", alignItems: "center",
        position: "relative", zIndex: 1, background: "rgba(2,8,24,0.6)"
      }}>
        <span style={{ color: "rgba(0,120,200,0.35)", fontSize: 11, fontFamily: "'DM Mono', monospace" }}>
          © 2025 Ashok Leyland Ltd. — Internal Operations Research Tool
        </span>
        <span style={{ color: "rgba(0,120,200,0.35)", fontSize: 10, fontFamily: "'DM Mono', monospace", letterSpacing: 1 }}>
          AHP · Entropy · TOPSIS · FLO
        </span>
      </footer>
    </div>
  );
}