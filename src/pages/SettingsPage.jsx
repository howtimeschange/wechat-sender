import { useState, useEffect, useRef } from 'react'
import './SettingsPage.css'

const DEFAULT_CFG = {
  send_interval: 5,
  max_per_minute: 8,
  poll_seconds: 15,
  dry_run: false,
}

export default function SettingsPage() {
  const [cfg, setCfg] = useState(DEFAULT_CFG)
  const [saved, setSaved] = useState(false)
  const [daemonOn, setDaemonOn] = useState(false)
  const logEndRef = useRef(null)
  const [daemonLogs, setDaemonLogs] = useState([])

  useEffect(() => {
    if (window.api) {
      window.api.getConfig().then(data => {
        if (data && !data.error) setCfg({ ...DEFAULT_CFG, ...data })
      })
      window.api.daemonStatus().then(({ running }) => setDaemonOn(running))

      const unsub1 = window.api.onDaemonLog(data => {
        const ts = new Date().toLocaleTimeString('zh-CN', { hour12: false })
        setDaemonLogs(prev => [...prev.slice(-300), { text: `[${ts}] ${data.trim()}`, id: Date.now() + Math.random() }])
      })
      const unsub2 = window.api.onDaemonStopped(() => {
        setDaemonOn(false)
        setDaemonLogs(prev => [...prev, { text: '— 守护进程已停止 —', id: Date.now() }])
      })
      return () => { unsub1?.(); unsub2?.() }
    }
  }, [])

  useEffect(() => {
    logEndRef.current?.scrollIntoView({ behavior: 'smooth' })
  }, [daemonLogs])

  const save = async () => {
    if (window.api) await window.api.saveConfig(cfg)
    setSaved(true)
    setTimeout(() => setSaved(false), 2500)
  }

  const set = (k, v) => { setCfg(c => ({ ...c, [k]: v })); setSaved(false) }

  // ── 守护进程 ──────────────────────────────
  const toggleDaemon = async () => {
    if (!window.api) { setDaemonOn(v => !v); return }
    if (daemonOn) {
      await window.api.daemonStop()
      setDaemonOn(false)
    } else {
      const r = await window.api.daemonStart()
      if (r.ok) { setDaemonOn(true); setDaemonLogs(prev => [...prev, { text: '▶ 守护进程启动中...', id: Date.now() }]) }
      else setDaemonLogs(prev => [...prev, { text: `❌ ${r.error}`, id: Date.now() }])
    }
  }

  return (
    <div className="settings-page page">
      <h1 className="page-title">⚙️ 设置</h1>

      {/* 发送参数 */}
      <div className="card settings-card">
        <div className="settings-section-title">发送参数</div>

        <div className="cfg-row">
          <div className="cfg-label">
            <span>发送间隔</span>
            <span className="cfg-hint">两条消息之间的等待时间（秒）</span>
          </div>
          <div className="cfg-value">
            <div className="cfg-number-row">
              <input type="range" min="1" max="30" step="1" value={cfg.send_interval}
                onChange={e => set('send_interval', Number(e.target.value))} className="cfg-slider" />
              <span className="cfg-number-badge">{cfg.send_interval}s</span>
            </div>
          </div>
        </div>

        <div className="cfg-row">
          <div className="cfg-label">
            <span>每分钟上限</span>
            <span className="cfg-hint">防止发送过快导致封号</span>
          </div>
          <div className="cfg-value">
            <div className="cfg-number-row">
              <input type="range" min="1" max="30" step="1" value={cfg.max_per_minute}
                onChange={e => set('max_per_minute', Number(e.target.value))} className="cfg-slider" />
              <span className="cfg-number-badge">{cfg.max_per_minute}条</span>
            </div>
          </div>
        </div>

        <div className="cfg-row">
          <div className="cfg-label">
            <span>轮询间隔</span>
            <span className="cfg-hint">守护进程检查新任务的频率（秒）</span>
          </div>
          <div className="cfg-value">
            <div className="cfg-number-row">
              <input type="range" min="5" max="120" step="5" value={cfg.poll_seconds}
                onChange={e => set('poll_seconds', Number(e.target.value))} className="cfg-slider" />
              <span className="cfg-number-badge">{cfg.poll_seconds}s</span>
            </div>
          </div>
        </div>

        <div className="cfg-row cfg-toggle-row">
          <div className="cfg-label">
            <span>模拟运行</span>
            <span className="cfg-hint">开启后不真实发送消息</span>
          </div>
          <label className="toggle">
            <input type="checkbox" checked={cfg.dry_run}
              onChange={e => set('dry_run', e.target.checked)} />
            <span className="toggle-slider" />
          </label>
        </div>

        {cfg.dry_run && (
          <div className="dry-run-warning">⚠️ 模拟运行模式开启，不会真实发送任何消息</div>
        )}
      </div>

      {/* 守护进程 */}
      <div className="card settings-card">
        <div className="settings-section-title">守护进程</div>

        <div className="daemon-row">
          <div className="daemon-status">
            <div className={`daemon-dot ${daemonOn ? 'dot-running' : 'dot-stopped'}`} />
            <div>
              <div className="daemon-status-label">{daemonOn ? '运行中' : '已停止'}</div>
              <div className="daemon-status-sub">
                {daemonOn
                  ? `每 ${cfg.poll_seconds}s 检查一次任务，到时立即发送`
                  : '监控任务发送时间，到点自动执行'}
              </div>
            </div>
          </div>
          <button className={`btn ${daemonOn ? 'btn-danger' : 'btn-primary'} daemon-toggle-btn`} onClick={toggleDaemon}>
            {daemonOn ? '⏹ 停止' : '▶ 启动'}
          </button>
        </div>

        {daemonLogs.length > 0 && (
          <div className="daemon-log-mini">
            {daemonLogs.slice(-6).map(l => (
              <div key={l.id} className="log-mini-line">{l.text}</div>
            ))}
            <div ref={logEndRef} />
          </div>
        )}
      </div>

      {/* 说明 */}
      <div className="card settings-card">
        <div className="settings-section-title">定时说明</div>
        <p className="settings-note">
          每个任务可设置「发送时间」，守护进程会在指定时间自动发送。<br/>
          轮询间隔决定检查频率，建议 15–30 秒。<br/>
          发送后任务状态变为「发送成功」，重复任务（repeat）会在下一周期重新生效。
        </p>
      </div>

      {/* 保存 */}
      <div className="save-row">
        <button className="btn btn-primary save-btn" onClick={save}>保存设置</button>
        {saved && <span className="saved-tip">✅ 已保存</span>}
      </div>
    </div>
  )
}
