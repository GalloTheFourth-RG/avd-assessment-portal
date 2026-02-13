import React, { useState, useEffect, useCallback, useRef } from 'react'
import { Monitor, Play, Clock, FileText, Download, AlertTriangle, CheckCircle, 
         Settings, ChevronRight, Loader2, RefreshCw, Shield, Cpu, DollarSign,
         Activity, Server, HardDrive, BarChart3, Eye } from 'lucide-react'

// ============================================================================
// Styles
// ============================================================================
const styles = `
  * { margin: 0; padding: 0; box-sizing: border-box; }
  
  :root {
    --bg-primary: #0a0e1a;
    --bg-secondary: #111827;
    --bg-card: #1a2035;
    --bg-card-hover: #1e2640;
    --bg-input: #0d1220;
    --border: #2a3550;
    --border-focus: #3b82f6;
    --text-primary: #e2e8f0;
    --text-secondary: #94a3b8;
    --text-muted: #64748b;
    --accent: #3b82f6;
    --accent-hover: #2563eb;
    --accent-glow: rgba(59, 130, 246, 0.15);
    --success: #10b981;
    --warning: #f59e0b;
    --danger: #ef4444;
    --surface-success: rgba(16, 185, 129, 0.1);
    --surface-warning: rgba(245, 158, 11, 0.1);
    --surface-danger: rgba(239, 68, 68, 0.1);
    --font-body: 'DM Sans', -apple-system, sans-serif;
    --font-mono: 'JetBrains Mono', monospace;
    --radius: 12px;
    --radius-sm: 8px;
  }
  
  body {
    font-family: var(--font-body);
    background: var(--bg-primary);
    color: var(--text-primary);
    min-height: 100vh;
    -webkit-font-smoothing: antialiased;
  }
  
  .app {
    display: flex;
    min-height: 100vh;
  }
  
  /* ── Sidebar ── */
  .sidebar {
    width: 260px;
    background: var(--bg-secondary);
    border-right: 1px solid var(--border);
    display: flex;
    flex-direction: column;
    padding: 24px 16px;
    gap: 8px;
    position: fixed;
    top: 0;
    left: 0;
    bottom: 0;
    z-index: 10;
  }
  
  .sidebar-brand {
    padding: 0 8px 20px;
    border-bottom: 1px solid var(--border);
    margin-bottom: 12px;
  }
  
  .sidebar-brand h1 {
    font-size: 15px;
    font-weight: 600;
    letter-spacing: -0.02em;
    color: var(--text-primary);
  }
  
  .sidebar-brand p {
    font-size: 11px;
    color: var(--text-muted);
    margin-top: 4px;
    font-family: var(--font-mono);
  }
  
  .nav-item {
    display: flex;
    align-items: center;
    gap: 10px;
    padding: 10px 12px;
    border-radius: var(--radius-sm);
    cursor: pointer;
    transition: all 0.15s;
    font-size: 13px;
    font-weight: 500;
    color: var(--text-secondary);
    border: 1px solid transparent;
  }
  
  .nav-item:hover { background: var(--bg-card); color: var(--text-primary); }
  .nav-item.active { 
    background: var(--accent-glow); 
    color: var(--accent); 
    border-color: rgba(59, 130, 246, 0.2);
  }
  
  .nav-item svg { width: 16px; height: 16px; opacity: 0.7; }
  .nav-item.active svg { opacity: 1; }
  
  .nav-section {
    font-size: 10px;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    color: var(--text-muted);
    padding: 16px 12px 6px;
    font-weight: 600;
  }
  
  /* ── Main Content ── */
  .main {
    flex: 1;
    margin-left: 260px;
    padding: 32px 40px;
    max-width: 1200px;
  }
  
  .page-header {
    margin-bottom: 32px;
  }
  
  .page-header h2 {
    font-size: 24px;
    font-weight: 600;
    letter-spacing: -0.02em;
  }
  
  .page-header p {
    font-size: 14px;
    color: var(--text-secondary);
    margin-top: 6px;
  }
  
  /* ── Cards ── */
  .card {
    background: var(--bg-card);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 24px;
    margin-bottom: 20px;
    transition: border-color 0.15s;
  }
  
  .card:hover { border-color: #3a4a6a; }
  
  .card-header {
    display: flex;
    align-items: center;
    gap: 10px;
    margin-bottom: 16px;
  }
  
  .card-header h3 {
    font-size: 14px;
    font-weight: 600;
    letter-spacing: -0.01em;
  }
  
  .card-header svg { width: 18px; height: 18px; color: var(--accent); }
  
  /* ── Form Elements ── */
  .form-group { margin-bottom: 20px; }
  
  .form-label {
    display: block;
    font-size: 12px;
    font-weight: 500;
    color: var(--text-secondary);
    margin-bottom: 6px;
    letter-spacing: 0.01em;
  }
  
  .form-input, .form-select {
    width: 100%;
    padding: 10px 14px;
    background: var(--bg-input);
    border: 1px solid var(--border);
    border-radius: var(--radius-sm);
    color: var(--text-primary);
    font-family: var(--font-mono);
    font-size: 13px;
    transition: border-color 0.15s;
    outline: none;
  }
  
  .form-input:focus, .form-select:focus { border-color: var(--border-focus); }
  .form-input::placeholder { color: var(--text-muted); }
  
  .form-hint {
    font-size: 11px;
    color: var(--text-muted);
    margin-top: 4px;
  }
  
  .form-row {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 16px;
  }
  
  .toggle-row {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 10px 0;
    border-bottom: 1px solid var(--border);
  }
  
  .toggle-row:last-child { border-bottom: none; }
  
  .toggle-label {
    font-size: 13px;
    font-weight: 500;
  }
  
  .toggle-desc {
    font-size: 11px;
    color: var(--text-muted);
    margin-top: 2px;
  }
  
  .toggle {
    width: 40px;
    height: 22px;
    background: var(--border);
    border-radius: 11px;
    position: relative;
    cursor: pointer;
    transition: background 0.2s;
    flex-shrink: 0;
    border: none;
  }
  
  .toggle.active { background: var(--accent); }
  
  .toggle::after {
    content: '';
    width: 16px;
    height: 16px;
    background: white;
    border-radius: 50%;
    position: absolute;
    top: 3px;
    left: 3px;
    transition: transform 0.2s;
  }
  
  .toggle.active::after { transform: translateX(18px); }
  
  /* ── Buttons ── */
  .btn {
    display: inline-flex;
    align-items: center;
    gap: 8px;
    padding: 10px 20px;
    border-radius: var(--radius-sm);
    font-size: 13px;
    font-weight: 600;
    font-family: var(--font-body);
    cursor: pointer;
    transition: all 0.15s;
    border: 1px solid transparent;
  }
  
  .btn-primary {
    background: var(--accent);
    color: white;
    border-color: var(--accent);
  }
  .btn-primary:hover { background: var(--accent-hover); }
  .btn-primary:disabled { opacity: 0.5; cursor: not-allowed; }
  
  .btn-secondary {
    background: transparent;
    color: var(--text-secondary);
    border-color: var(--border);
  }
  .btn-secondary:hover { border-color: var(--text-muted); color: var(--text-primary); }
  
  .btn svg { width: 16px; height: 16px; }
  
  /* ── Status / Progress ── */
  .status-bar {
    background: var(--bg-secondary);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 20px 24px;
    margin-bottom: 20px;
  }
  
  .status-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-bottom: 12px;
  }
  
  .status-title {
    display: flex;
    align-items: center;
    gap: 10px;
    font-size: 14px;
    font-weight: 600;
  }
  
  .progress-bar {
    height: 4px;
    background: var(--border);
    border-radius: 2px;
    overflow: hidden;
  }
  
  .progress-fill {
    height: 100%;
    background: var(--accent);
    border-radius: 2px;
    transition: width 0.5s ease;
  }
  
  .progress-fill.indeterminate {
    width: 30%;
    animation: indeterminate 1.5s infinite;
  }
  
  @keyframes indeterminate {
    0% { transform: translateX(-100%); }
    100% { transform: translateX(400%); }
  }
  
  @keyframes spin {
    to { transform: rotate(360deg); }
  }
  
  .spinner { animation: spin 1s linear infinite; }
  
  /* ── Results / File List ── */
  .file-list { list-style: none; }
  
  .file-item {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 12px 16px;
    border-bottom: 1px solid var(--border);
    transition: background 0.1s;
  }
  
  .file-item:hover { background: var(--bg-card-hover); }
  .file-item:last-child { border-bottom: none; }
  
  .file-name {
    display: flex;
    align-items: center;
    gap: 10px;
    font-size: 13px;
    font-family: var(--font-mono);
  }
  
  .file-size {
    font-size: 11px;
    color: var(--text-muted);
    font-family: var(--font-mono);
  }
  
  .file-actions {
    display: flex;
    gap: 8px;
  }
  
  .file-btn {
    padding: 4px 10px;
    font-size: 11px;
    font-weight: 500;
    border-radius: 6px;
    border: 1px solid var(--border);
    background: transparent;
    color: var(--text-secondary);
    cursor: pointer;
    font-family: var(--font-body);
    display: flex;
    align-items: center;
    gap: 4px;
    transition: all 0.15s;
  }
  
  .file-btn:hover { border-color: var(--accent); color: var(--accent); }
  .file-btn svg { width: 12px; height: 12px; }
  
  /* ── Report Viewer ── */
  .report-frame {
    width: 100%;
    height: calc(100vh - 180px);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    background: white;
  }
  
  /* ── Subscription Chips ── */
  .sub-chips {
    display: flex;
    flex-wrap: wrap;
    gap: 8px;
    margin-top: 8px;
  }
  
  .sub-chip {
    display: flex;
    align-items: center;
    gap: 6px;
    padding: 6px 12px;
    border-radius: 20px;
    font-size: 12px;
    cursor: pointer;
    transition: all 0.15s;
    border: 1px solid var(--border);
    background: var(--bg-input);
    color: var(--text-secondary);
  }
  
  .sub-chip.selected {
    background: var(--accent-glow);
    border-color: var(--accent);
    color: var(--accent);
  }
  
  .sub-chip:hover { border-color: var(--text-muted); }
  
  /* ── Run History ── */
  .run-item {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 14px 16px;
    border-bottom: 1px solid var(--border);
    cursor: pointer;
    transition: background 0.1s;
  }
  
  .run-item:hover { background: var(--bg-card-hover); }
  
  .run-id {
    font-family: var(--font-mono);
    font-size: 13px;
  }
  
  .run-meta {
    font-size: 11px;
    color: var(--text-muted);
  }
  
  /* ── Empty State ── */
  .empty-state {
    text-align: center;
    padding: 60px 20px;
    color: var(--text-muted);
  }
  
  .empty-state svg {
    width: 48px;
    height: 48px;
    margin-bottom: 16px;
    opacity: 0.3;
  }
  
  .empty-state h3 {
    font-size: 16px;
    font-weight: 500;
    color: var(--text-secondary);
    margin-bottom: 6px;
  }
  
  .empty-state p { font-size: 13px; }
  
  /* ── Badge ── */
  .badge {
    font-size: 10px;
    font-weight: 600;
    padding: 2px 8px;
    border-radius: 10px;
    text-transform: uppercase;
    letter-spacing: 0.04em;
  }
  
  .badge-success { background: var(--surface-success); color: var(--success); }
  .badge-warning { background: var(--surface-warning); color: var(--warning); }
  .badge-danger { background: var(--surface-danger); color: var(--danger); }
`

// ============================================================================
// API helpers
// ============================================================================
const api = {
  health: () => fetch('/api/health').then(r => r.json()),
  subscriptions: () => fetch('/api/subscriptions').then(r => r.json()),
  startAssessment: (config) => fetch('/api/assess', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(config)
  }).then(r => r.json()),
  assessmentStatus: (runId) => fetch(`/api/assess/${runId}`).then(r => r.json()),
  listResults: (runId) => fetch(`/api/results/${runId}`).then(r => r.json()),
  listRuns: () => fetch('/api/runs').then(r => r.json()),
  deleteRun: (runId) => fetch(`/api/runs/${runId}`, { method: 'DELETE' }).then(r => r.json()),
}

function formatBytes(bytes) {
  if (!bytes) return '0 B'
  const k = 1024
  const sizes = ['B', 'KB', 'MB', 'GB']
  const i = Math.floor(Math.log(bytes) / Math.log(k))
  return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i]
}

function formatDuration(seconds) {
  if (seconds < 60) return `${seconds}s`
  const m = Math.floor(seconds / 60)
  const s = seconds % 60
  return `${m}m ${s}s`
}

// ============================================================================
// App Component
// ============================================================================
export default function App() {
  const [page, setPage] = useState('assess')
  const [health, setHealth] = useState(null)
  const [subscriptions, setSubs] = useState([])
  const [selectedSubs, setSelectedSubs] = useState([])
  const [runs, setRuns] = useState([])
  
  // Assessment config
  const [tenantId, setTenantId] = useState('')
  const [laWorkspaces, setLaWorkspaces] = useState('')
  const [lookbackDays, setLookbackDays] = useState(7)
  const [includeAdvisor, setIncludeAdvisor] = useState(true)
  const [includeReservations, setIncludeReservations] = useState(false)
  const [skipCosts, setSkipCosts] = useState(false)
  const [scrubPII, setScrubPII] = useState(false)
  const [quickSummary, setQuickSummary] = useState(false)
  const [companyName, setCompanyName] = useState('')
  const [analystName, setAnalystName] = useState('')
  
  // Run state
  const [activeRun, setActiveRun] = useState(null)
  const [runStatus, setRunStatus] = useState(null)
  const [results, setResults] = useState(null)
  const [viewingReport, setViewingReport] = useState(null)
  
  const pollRef = useRef(null)
  
  // Init: check health & load subs
  useEffect(() => {
    api.health().then(h => { 
      setHealth(h)
      if (h.tenantId && !tenantId) setTenantId(h.tenantId)
    }).catch(() => setHealth({ status: 'error' }))
    api.subscriptions()
      .then(d => { if (d.subscriptions) setSubs(d.subscriptions) })
      .catch(() => {})
    api.listRuns()
      .then(d => { if (d.runs) setRuns(d.runs) })
      .catch(() => {})
  }, [])
  
  // Poll active run
  useEffect(() => {
    if (!activeRun || activeRun === 'running') return
    
    pollRef.current = setInterval(async () => {
      try {
        const status = await api.assessmentStatus(activeRun)
        setRunStatus(status)
        
        if (status.status === 'completed' || status.status === 'failed') {
          clearInterval(pollRef.current)
          if (status.status === 'completed') {
            const res = await api.listResults(activeRun)
            setResults(res)
          }
          api.listRuns().then(d => { if (d.runs) setRuns(d.runs) }).catch(() => {})
        }
      } catch (e) {
        console.error('Poll error:', e)
      }
    }, 3000)
    
    return () => clearInterval(pollRef.current)
  }, [activeRun])
  
  const toggleSub = (subId) => {
    setSelectedSubs(prev => 
      prev.includes(subId) ? prev.filter(s => s !== subId) : [...prev, subId]
    )
  }
  
  const startAssessment = async () => {
    const config = {
      tenantId,
      subscriptionIds: selectedSubs,
      logAnalyticsWorkspaceIds: laWorkspaces ? laWorkspaces.split(',').map(s => s.trim()).filter(Boolean) : [],
      metricsLookbackDays: lookbackDays,
      includeAdvisor,
      includeReservations,
      skipCosts,
      scrubPII,
      quickSummary,
      companyName,
      analystName,
    }
    
    try {
      const res = await api.startAssessment(config)
      setActiveRun(res.runId)
      setRunStatus({ status: 'started', elapsedSeconds: 0 })
      setResults(null)
      setPage('status')
    } catch (e) {
      console.error('Start failed:', e)
    }
  }
  
  const viewRun = async (runId) => {
    setActiveRun(runId)
    try {
      const res = await api.listResults(runId)
      setResults(res)
      setRunStatus({ status: 'completed' })
      setPage('status')
    } catch (e) {
      console.error('Load results failed:', e)
    }
  }
  
  const deleteRun = async (e, runId) => {
    e.stopPropagation()
    if (!confirm(`Delete run ${runId} and all its files?`)) return
    try {
      await api.deleteRun(runId)
      setRuns(prev => prev.filter(r => r.runId !== runId))
      if (activeRun === runId) {
        setActiveRun(null)
        setResults(null)
        setRunStatus(null)
      }
    } catch (err) {
      console.error('Delete failed:', err)
    }
  }
  
  const canStart = tenantId && selectedSubs.length > 0 && !activeRun
  
  // ── Render ──
  return (
    <>
      <style>{styles}</style>
      <div className="app">
        {/* Sidebar */}
        <nav className="sidebar">
          <div className="sidebar-brand">
            <h1>AVD Assessment</h1>
            <p>v4.1.0 · Portal v2.0</p>
          </div>
          
          <div className="nav-section">Assessment</div>
          <div className={`nav-item ${page === 'assess' ? 'active' : ''}`} onClick={() => setPage('assess')}>
            <Play /> New Assessment
          </div>
          <div className={`nav-item ${page === 'status' ? 'active' : ''}`} onClick={() => setPage('status')}>
            <Activity /> {activeRun ? 'Current Run' : 'Results'}
          </div>
          
          <div className="nav-section">History</div>
          <div className={`nav-item ${page === 'history' ? 'active' : ''}`} onClick={() => setPage('history')}>
            <Clock /> Past Runs
          </div>
          
          <div className="nav-section">System</div>
          <div className={`nav-item ${page === 'settings' ? 'active' : ''}`} onClick={() => setPage('settings')}>
            <Settings /> Configuration
          </div>
          
          {/* Health indicator */}
          <div style={{ marginTop: 'auto', padding: '12px', borderTop: '1px solid var(--border)' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 11, color: 'var(--text-muted)' }}>
              <div style={{ 
                width: 6, height: 6, borderRadius: '50%', 
                background: health?.status === 'healthy' ? 'var(--success)' : 'var(--danger)' 
              }} />
              {health?.status === 'healthy' ? 'System healthy' : 'Connecting...'}
            </div>
          </div>
        </nav>
        
        {/* Main Content */}
        <main className="main">
          
          {/* ── PAGE: New Assessment ── */}
          {page === 'assess' && (
            <>
              <div className="page-header">
                <h2>New Assessment</h2>
                <p>Configure and launch an AVD environment assessment</p>
              </div>
              
              {/* Target */}
              <div className="card">
                <div className="card-header">
                  <Server /> <h3>Target Environment</h3>
                </div>
                
                <div className="form-group">
                  <label className="form-label">Tenant ID {health?.tenantId ? <span style={{color:'var(--success)', fontWeight: 400}}>(auto-detected)</span> : ''}</label>
                  <input 
                    className="form-input" 
                    placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
                    value={tenantId}
                    onChange={e => setTenantId(e.target.value)}
                  />
                </div>
                
                <div className="form-group">
                  <label className="form-label">Subscriptions</label>
                  {subscriptions.length > 0 ? (
                    <div className="sub-chips">
                      {subscriptions.map(sub => (
                        <div 
                          key={sub.id}
                          className={`sub-chip ${selectedSubs.includes(sub.id) ? 'selected' : ''}`}
                          onClick={() => toggleSub(sub.id)}
                        >
                          {selectedSubs.includes(sub.id) ? <CheckCircle size={12} /> : null}
                          {sub.name}
                        </div>
                      ))}
                    </div>
                  ) : (
                    <p className="form-hint">
                      No subscriptions loaded. Enter a tenant ID and ensure the managed identity has Reader access.
                    </p>
                  )}
                </div>
                
                <div className="form-group">
                  <label className="form-label">Log Analytics Workspace IDs (comma-separated, optional)</label>
                  <input 
                    className="form-input"
                    placeholder="/subscriptions/.../workspaces/your-workspace"
                    value={laWorkspaces}
                    onChange={e => setLaWorkspaces(e.target.value)}
                  />
                  <p className="form-hint">Required for session analytics, UX scoring, login times, and user counts</p>
                </div>
              </div>
              
              {/* Options */}
              <div className="card">
                <div className="card-header">
                  <Settings /> <h3>Analysis Options</h3>
                </div>
                
                <div className="form-row" style={{ marginBottom: 16 }}>
                  <div className="form-group">
                    <label className="form-label">Metrics Lookback (days)</label>
                    <input 
                      className="form-input" type="number" min={1} max={90}
                      value={lookbackDays}
                      onChange={e => setLookbackDays(parseInt(e.target.value) || 7)}
                    />
                  </div>
                  <div className="form-group">
                    <label className="form-label">Company Name (optional)</label>
                    <input className="form-input" placeholder="Contoso Ltd" value={companyName} onChange={e => setCompanyName(e.target.value)} />
                  </div>
                </div>
                
                <div className="toggle-row">
                  <div>
                    <div className="toggle-label">Azure Advisor</div>
                    <div className="toggle-desc">Include Microsoft Advisor recommendations</div>
                  </div>
                  <button className={`toggle ${includeAdvisor ? 'active' : ''}`} onClick={() => setIncludeAdvisor(!includeAdvisor)} />
                </div>
                
                <div className="toggle-row">
                  <div>
                    <div className="toggle-label">Reservation Analysis</div>
                    <div className="toggle-desc">Analyze RI coverage and savings (requires Az.Reservations)</div>
                  </div>
                  <button className={`toggle ${includeReservations ? 'active' : ''}`} onClick={() => setIncludeReservations(!includeReservations)} />
                </div>
                
                <div className="toggle-row">
                  <div>
                    <div className="toggle-label">Scrub PII</div>
                    <div className="toggle-desc">Anonymize usernames, VM names, IPs in reports</div>
                  </div>
                  <button className={`toggle ${scrubPII ? 'active' : ''}`} onClick={() => setScrubPII(!scrubPII)} />
                </div>
                
                <div className="toggle-row">
                  <div>
                    <div className="toggle-label">Skip Cost Analysis</div>
                    <div className="toggle-desc">Skip Cost Management API (faster, uses PAYG estimates)</div>
                  </div>
                  <button className={`toggle ${skipCosts ? 'active' : ''}`} onClick={() => setSkipCosts(!skipCosts)} />
                </div>
                
                <div className="toggle-row">
                  <div>
                    <div className="toggle-label">Quick Summary Only</div>
                    <div className="toggle-desc">2-minute config check without metrics (great for triage)</div>
                  </div>
                  <button className={`toggle ${quickSummary ? 'active' : ''}`} onClick={() => setQuickSummary(!quickSummary)} />
                </div>
              </div>
              
              {/* Launch */}
              <div style={{ display: 'flex', gap: 12, alignItems: 'center' }}>
                <button 
                  className="btn btn-primary" 
                  disabled={!canStart}
                  onClick={startAssessment}
                >
                  <Play size={16} />
                  {quickSummary ? 'Run Quick Summary' : 'Start Full Assessment'}
                </button>
                <span style={{ fontSize: 12, color: 'var(--text-muted)' }}>
                  {quickSummary ? 'Est. 2-3 minutes' : `Est. 15-45 min depending on environment size`}
                </span>
              </div>
            </>
          )}
          
          {/* ── PAGE: Status / Results ── */}
          {page === 'status' && (
            <>
              <div className="page-header">
                <h2>{runStatus?.status === 'completed' ? 'Assessment Results' : 'Assessment in Progress'}</h2>
                {activeRun && <p style={{ fontFamily: 'var(--font-mono)', fontSize: 12 }}>{activeRun}</p>}
              </div>
              
              {/* Progress */}
              {runStatus && runStatus.status !== 'completed' && (
                <div className="status-bar">
                  <div className="status-header">
                    <div className="status-title">
                      {runStatus.status === 'running' || runStatus.status === 'started' ? (
                        <><Loader2 className="spinner" size={18} /> Running assessment...</>
                      ) : runStatus.status === 'failed' ? (
                        <><AlertTriangle size={18} style={{color:'var(--danger)'}}/> Assessment failed</>
                      ) : (
                        <><Clock size={18} /> {runStatus.status}</>
                      )}
                    </div>
                    <span style={{ fontSize: 12, color: 'var(--text-muted)', fontFamily: 'var(--font-mono)' }}>
                      {formatDuration(runStatus.elapsedSeconds || 0)}
                    </span>
                  </div>
                  <div className="progress-bar">
                    <div className={`progress-fill ${runStatus.status === 'running' ? 'indeterminate' : ''}`} />
                  </div>
                  {runStatus.error && (
                    <p style={{ marginTop: 12, fontSize: 13, color: 'var(--danger)', fontFamily: 'var(--font-mono)' }}>
                      {runStatus.error}
                    </p>
                  )}
                </div>
              )}
              
              {/* Completed: show results */}
              {runStatus?.status === 'completed' && results && (
                <>
                  <div className="status-bar" style={{ borderColor: 'rgba(16, 185, 129, 0.3)', background: 'rgba(16, 185, 129, 0.05)' }}>
                    <div className="status-title" style={{ color: 'var(--success)' }}>
                      <CheckCircle size={18} /> Assessment complete — {results.files?.length || 0} files generated
                    </div>
                  </div>
                  
                  {/* View HTML Report button */}
                  {results.files?.some(f => f.name.endsWith('.html')) && !viewingReport && (
                    <button 
                      className="btn btn-primary" 
                      style={{ marginBottom: 20 }}
                      onClick={() => {
                        const html = results.files.find(f => f.name.endsWith('.html'))
                        if (html) setViewingReport(html.url)
                      }}
                    >
                      <Eye size={16} /> View HTML Dashboard
                    </button>
                  )}
                  
                  {/* Embedded report */}
                  {viewingReport && (
                    <div style={{ marginBottom: 20 }}>
                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
                        <span style={{ fontSize: 13, fontWeight: 600 }}>Interactive Dashboard</span>
                        <button className="btn btn-secondary" onClick={() => setViewingReport(null)} style={{ padding: '6px 12px', fontSize: 12 }}>
                          Close
                        </button>
                      </div>
                      <iframe src={viewingReport} className="report-frame" title="AVD Assessment Report" />
                    </div>
                  )}
                  
                  {/* File list */}
                  <div className="card">
                    <div className="card-header">
                      <FileText /> <h3>Output Files</h3>
                    </div>
                    <ul className="file-list">
                      {results.files?.map((file, i) => (
                        <li key={i} className="file-item">
                          <div className="file-name">
                            <FileText size={14} style={{ opacity: 0.5 }} />
                            {file.name}
                          </div>
                          <div style={{ display: 'flex', alignItems: 'center', gap: 16 }}>
                            <span className="file-size">{formatBytes(file.size)}</span>
                            <div className="file-actions">
                              {file.name.endsWith('.html') && (
                                <button className="file-btn" onClick={() => setViewingReport(file.url)}>
                                  <Eye /> View
                                </button>
                              )}
                              <a className="file-btn" href={file.url} download style={{ textDecoration: 'none' }}>
                                <Download /> Download
                              </a>
                            </div>
                          </div>
                        </li>
                      ))}
                    </ul>
                  </div>
                </>
              )}
              
              {/* No active run */}
              {!activeRun && (
                <div className="empty-state">
                  <BarChart3 />
                  <h3>No active assessment</h3>
                  <p>Start a new assessment from the sidebar to see results here.</p>
                </div>
              )}
            </>
          )}
          
          {/* ── PAGE: History ── */}
          {page === 'history' && (
            <>
              <div className="page-header">
                <h2>Past Runs</h2>
                <p>View and download results from previous assessments</p>
              </div>
              
              <div className="card" style={{ padding: 0 }}>
                {runs.length > 0 ? (
                  runs.map((run, i) => (
                    <div key={i} className="run-item" onClick={() => viewRun(run.runId)}>
                      <div>
                        <div className="run-id">{run.runId}</div>
                        <div className="run-meta">
                          {run.files} files · {formatBytes(run.totalSize)}
                        </div>
                      </div>
                      <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <button 
                          onClick={(e) => deleteRun(e, run.runId)}
                          style={{ 
                            background: 'transparent', border: '1px solid var(--danger)', 
                            color: 'var(--danger)', borderRadius: '6px', padding: '4px 10px',
                            cursor: 'pointer', fontSize: '12px', opacity: 0.7,
                            transition: 'opacity 0.2s'
                          }}
                          onMouseEnter={e => e.target.style.opacity = 1}
                          onMouseLeave={e => e.target.style.opacity = 0.7}
                        >Delete</button>
                        <ChevronRight size={16} style={{ color: 'var(--text-muted)' }} />
                      </div>
                    </div>
                  ))
                ) : (
                  <div className="empty-state">
                    <Clock />
                    <h3>No past runs</h3>
                    <p>Completed assessments will appear here.</p>
                  </div>
                )}
              </div>
            </>
          )}
          
          {/* ── PAGE: Settings ── */}
          {page === 'settings' && (
            <>
              <div className="page-header">
                <h2>Configuration</h2>
                <p>System status and portal settings</p>
              </div>
              
              <div className="card">
                <div className="card-header">
                  <Shield /> <h3>System Status</h3>
                </div>
                
                <div style={{ display: 'grid', gap: 12 }}>
                  {[
                    { label: 'Backend', ok: health?.status === 'healthy', detail: health?.status || 'unknown' },
                    { label: 'Assessment Script', ok: health?.scriptExists, detail: health?.scriptExists ? 'Found' : 'Missing' },
                    { label: 'Storage Account', ok: health?.storageConfigured, detail: health?.storageConfigured ? 'Configured' : 'Not set' },
                    { label: 'Managed Identity', ok: health?.identityConfigured, detail: health?.identityConfigured ? 'Configured' : 'Not set' },
                  ].map((item, i) => (
                    <div key={i} style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '8px 0', borderBottom: '1px solid var(--border)' }}>
                      <span style={{ fontSize: 13 }}>{item.label}</span>
                      <span className={`badge ${item.ok ? 'badge-success' : 'badge-danger'}`}>
                        {item.detail}
                      </span>
                    </div>
                  ))}
                </div>
              </div>
              
              <div className="card">
                <div className="card-header">
                  <Monitor /> <h3>About</h3>
                </div>
                <p style={{ fontSize: 13, color: 'var(--text-secondary)', lineHeight: 1.7 }}>
                  AVD Assessment Portal runs the Enhanced AVD Evidence Pack script inside your Azure tenant 
                  using a managed identity. No credentials leave your environment. Assessment results are 
                  stored in your Storage Account and rendered in-browser.
                </p>
                <p style={{ fontSize: 12, color: 'var(--text-muted)', marginTop: 12, fontFamily: 'var(--font-mono)' }}>
                  Script v4.1.0 · Portal v1.0.0
                </p>
              </div>
            </>
          )}
        </main>
      </div>
    </>
  )
}
