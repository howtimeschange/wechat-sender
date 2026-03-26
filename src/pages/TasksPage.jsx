// BUILD_TEST_MARKER_1774415207
import { useState, useRef, useCallback, useEffect } from 'react'
import NewTaskModal from '../components/NewTaskModal'
import SendLogDrawer from '../components/SendLogDrawer'
import './TasksPage.css'

// ── 工具函数 ─────────────────────────────────────────────
let _uid = 1
const uid = () => String(_uid++)

const STATUS_LABELS = {
  waiting: '待发送',
  sending: '发送中',
  success: '发送成功',
  failed: '发送失败',
  stopped: '已停止',
}

const REPEAT_OPTIONS = ['', 'daily', 'weekly', 'workday']

// ── 解析 Excel（浏览器端，通过 IPC） ─────────────────────
async function parseExcel(file) {
  // 在 Electron 环境下通过 IPC 解析；浏览器 demo 下返回 mock
  if (window.api?.parseExcel) {
    const buf = await file.arrayBuffer()
    return window.api.parseExcel(Array.from(new Uint8Array(buf)))
  }
  // Mock 数据（开发预览用）
  return [
    { id: uid(), app: '微信', target: '张三', msg_type: '文字', text: '你好！这是测试消息', image_path: '', send_time: '', repeat: '', status: 'waiting' },
    { id: uid(), app: '微信', target: '产品讨论群', msg_type: '文字+图片', text: '活动图片来了', image_path: '/Users/demo/pic.png', send_time: '2026-03-26 09:00', repeat: '', status: 'waiting' },
    { id: uid(), app: '微信', target: '李四', msg_type: '文字', text: '明天会议记得参加', image_path: '', send_time: '', repeat: 'daily', status: 'waiting' },
  ]
}

const STORAGE_KEY = 'wechat_sender_tasks'

function loadTasks() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY)
    return raw ? JSON.parse(raw) : []
  } catch { return [] }
}
function saveTasks(tasks) {
  try { localStorage.setItem(STORAGE_KEY, JSON.stringify(tasks)) } catch {}
}

// ── 主组件 ───────────────────────────────────────────────
export default function TasksPage() {
  const [tasks, setTasks] = useState(loadTasks)
  const [selected, setSelected] = useState(new Set())
  const [sending, setSending] = useState(false)
  const [showNewTask, setShowNewTask] = useState(false)
  const [showLog, setShowLog] = useState(false)
  const [logs, setLogs] = useState([])
  const [dragging, setDragging] = useState(false)
  const [editTask, setEditTask] = useState(null)
  const [daemonOn, setDaemonOn] = useState(false)

  // ── 守护进程状态监听 ────────────────────────────────────
  useEffect(() => {
    if (!window.api) return
    window.api.daemonStatus().then(({ running }) => setDaemonOn(running))
    const unsubLog = window.api.onDaemonLog(data => {
      const ts = new Date().toLocaleTimeString('zh-CN', { hour12: false })
      setLogs(prev => [...prev.slice(-300), { text: `[${ts}] ${data.trim()}`, type: 'info', id: Date.now() + Math.random() }])
    })
    const unsubStop = window.api.onDaemonStopped(() => {
      setDaemonOn(false)
      setLogs(prev => [...prev, { text: `— 守护进程已停止 —`, type: 'info', id: Date.now() }])
    })
    return () => { unsubLog?.(); unsubStop?.() }
  }, [])

  // 持久化任务
  useEffect(() => { saveTasks(tasks) }, [tasks])
  const fileRef = useRef()
  const abortRef = useRef(false)

  // 统计
  const stats = {
    total: tasks.length,
    waiting: tasks.filter(t => t.status === 'waiting').length,
    success: tasks.filter(t => t.status === 'success').length,
    failed: tasks.filter(t => t.status === 'failed').length,
  }

  // ── 文件处理 ──────────────────────────────────────────
  const handleFile = async (file) => {
    if (!file) return
    if (!file.name.match(/\.(xlsx|xls)$/i)) {
      alert('请上传 Excel 文件（.xlsx / .xls）')
      return
    }
    try {
      const result = await parseExcel(file)
      if (!Array.isArray(result)) {
        alert('解析失败：' + (result.error || '未知错误'))
        return
      }
      const newTasks = result.map(r => ({ ...r, id: uid(), status: 'waiting' }))
      const allTasks = [...tasks, ...newTasks]
      window.api.saveGuiTasks(allTasks)
      setTasks(allTasks)
      setSelected(new Set())
    } catch (e) {
      alert('解析失败：' + e.message)
    }
  }

  const onDrop = useCallback((e) => {
    e.preventDefault()
    setDragging(false)
    const file = e.dataTransfer.files[0]
    handleFile(file)
  }, [])

  const onDragOver = (e) => { e.preventDefault(); setDragging(true) }
  const onDragLeave = () => setDragging(false)

  const downloadTemplate = () => {
    if (window.api?.downloadTemplate) {
      window.api.downloadTemplate()
    } else {
      // 开发模式：提示路径
      alert('模版将保存到桌面 wechat_template.xlsx')
    }
  }

  // ── 选择逻辑 ──────────────────────────────────────────
  const sendableTasks = tasks.filter(t => t.status === 'waiting' || t.status === 'failed' || t.status === 'stopped')
  const allSelected = sendableTasks.length > 0 && sendableTasks.every(t => selected.has(t.id))
  const someSelected = sendableTasks.some(t => selected.has(t.id))

  const toggleAll = () => {
    if (allSelected) {
      setSelected(new Set())
    } else {
      setSelected(new Set(sendableTasks.map(t => t.id)))
    }
  }

  const toggleOne = (id, currentlyChecked) => {
    setSelected(prev => {
      const s = new Set(prev)
      if (currentlyChecked) s.add(id)
      else s.delete(id)
      return s
    })
  }

  // ── 新建任务 ──────────────────────────────────────────
  const handleNewTask = (task) => {
    if (editTask) {
      setTasks(prev => {
        const updated = prev.map(t => t.id === editTask.id ? { ...task, id: editTask.id } : t)
        window.api?.saveGuiTasks(updated)
        return updated
      })
      setEditTask(null)
    } else {
      setTasks(prev => {
        const updated = [...prev, { ...task, id: uid(), status: 'waiting' }]
        window.api?.saveGuiTasks(updated)
        return updated
      })
    }
    setShowNewTask(false)
  }

  const deleteSelected = () => {
    if (!someSelected) return
    if (!confirm(`确认删除选中的 ${selected.size} 条任务？`)) return
    setTasks(prev => {
      const updated = prev.filter(t => !selected.has(t.id))
      window.api?.saveGuiTasks(updated)
      return updated
    })
    setSelected(new Set())
  }

  const deleteTask = (id) => {
    setTasks(prev => {
      const updated = prev.filter(t => t.id !== id)
      window.api?.saveGuiTasks(updated)
      return updated
    })
    setSelected(prev => { const s = new Set(prev); s.delete(id); return s })
  }

  // ── 发送逻辑 ──────────────────────────────────────────
  const addLog = (text, type = '') => {
    setLogs(prev => {
      const ts = new Date().toLocaleTimeString('zh-CN', { hour12: false })
      return [...prev.slice(-500), { text: `[${ts}] ${text}`, type, id: Date.now() + Math.random() }]
    })
  }

  const startSend = async () => {
    const targets = tasks.filter(t => selected.has(t.id) && (t.status === 'waiting' || t.status === 'failed' || t.status === 'stopped'))
    if (targets.length === 0) {
      alert('请先勾选要发送的任务')
      return
    }

    setSending(true)
    abortRef.current = false
    setShowLog(true)
    setLogs([])
    addLog(`▶ 开始发送 ${targets.length} 条任务...`)

    if (window.api) {
      // 真实 Electron 模式
      const unsub1 = window.api.onSendProgress((data) => {
        data.split('\n').filter(Boolean).forEach(line => {
          const type = line.includes('✅') ? 'success' : line.includes('❌') ? 'error' : ''
          addLog(line.trim(), type)
        })
      })
      const unsub3 = window.api.onTaskStatusUpdate(({ target, msg_type, status, error }) => {
        setTasks(prev => {
          const updated = prev.map(t =>
            t.target === target && t.msg_type === msg_type
              ? { ...t, status: status === 'success' ? 'success' : 'failed' }
              : t
          )
          window.api.saveGuiTasks(updated)
          return updated
        })
      })
      const unsub2 = window.api.onSendDone(() => {
        setSending(false)
        addLog('— 发送完成 —', 'success')
        unsub1?.(); unsub2?.(); unsub3?.()
      })
      window.api.sendSelected(targets)
    } else {
      // Dev mock
      for (const task of targets) {
        if (abortRef.current) {
          setTasks(prev => prev.map(t => targets.includes(t) && t.status === 'sending' ? { ...t, status: 'stopped' } : t))
          addLog('⏹ 已停止', 'error')
          break
        }
        setTasks(prev => prev.map(t => t.id === task.id ? { ...t, status: 'sending' } : t))
        addLog(`→ ${task.target} [${task.msg_type}]`)
        await new Promise(r => setTimeout(r, 800))

        const ok = Math.random() > 0.2
        setTasks(prev => prev.map(t => t.id === task.id ? { ...t, status: ok ? 'success' : 'failed' } : t))
        addLog(ok ? `  ✅ ${task.target} 发送成功` : `  ❌ ${task.target} 发送失败`, ok ? 'success' : 'error')
      }
      setSending(false)
      addLog('— 完成 —', 'success')
    }
  }

  const stopSend = () => {
    abortRef.current = true
    if (window.api?.stopSend) window.api.stopSend()
    setSending(false)
    addLog('⏹ 用户已停止发送', 'error')
  }

  const refreshStatus = async () => {
    // 可以通过 IPC 重新读取 Excel 状态
  }

  const clearCompleted = () => {
    setTasks(prev => {
      const updated = prev.filter(t => t.status !== 'success')
      window.api?.saveGuiTasks(updated)
      return updated
    })
    setSelected(prev => {
      const kept = new Set(tasks.filter(t => t.status !== 'success').map(t => t.id))
      return new Set([...prev].filter(id => kept.has(id)))
    })
  }

  return (
    <div className="tasks-page">
      {/* 顶栏 */}
      <div className="tasks-header">
        <div className="tasks-title-row">
          <h1 className="page-title">📤 发送任务</h1>
          {/* 守护进程开关（默认关闭，有 send_time 的任务才会自动发送） */}
          <div className="daemon-toggle-row">
            <span className={`daemon-dot ${daemonOn ? 'dot-running' : 'dot-stopped'}`} />
            <span className="daemon-label">{daemonOn ? '守护进程开启' : '守护进程关闭'}</span>
            <label className="toggle toggle-sm">
              <input type="checkbox" checked={daemonOn}
                onChange={async () => {
                  if (!window.api) { setDaemonOn(v => !v); return }
                  if (daemonOn) {
                    await window.api.daemonStop()
                    setDaemonOn(false)
                  } else {
                    await window.api.daemonStart()
                    setDaemonOn(true)
                  }
                }} />
              <span className="toggle-slider" />
            </label>
          </div>
          <div className="header-stats">
            <span className="stat-chip chip-total">共 {stats.total} 条</span>
            {stats.waiting > 0 && <span className="stat-chip chip-wait">待发 {stats.waiting}</span>}
            {stats.success > 0 && <span className="stat-chip chip-ok">成功 {stats.success}</span>}
            {stats.failed > 0 && <span className="stat-chip chip-fail">失败 {stats.failed}</span>}
          </div>
        </div>

        {/* 操作栏 */}
        <div className="tasks-toolbar">
          <div className="toolbar-left">
            {someSelected && (
              <>
                <span className="selected-hint">已选 {selected.size} 条</span>
                {!sending ? (
                  <button className="btn btn-primary btn-sm" onClick={startSend}>
                    ▶ 发送选中
                  </button>
                ) : (
                  <button className="btn btn-danger btn-sm" onClick={stopSend}>
                    ⏹ 停止发送
                  </button>
                )}
                <button className="btn btn-secondary btn-sm" onClick={deleteSelected}>
                  🗑 删除选中
                </button>
              </>
            )}
            {!someSelected && !sending && (
              <span className="hint-text">勾选任务后可批量发送</span>
            )}
          </div>
          <div className="toolbar-right">
            {stats.success > 0 && (
              <button className="btn btn-secondary btn-sm" onClick={clearCompleted}>
                清除已完成
              </button>
            )}
            {logs.length > 0 && (
              <button className="btn btn-ghost btn-sm" onClick={() => setShowLog(v => !v)}>
                {showLog ? '收起日志' : '查看日志'}
              </button>
            )}
            <button className="btn btn-secondary btn-sm" onClick={() => { setEditTask(null); setShowNewTask(true) }}>
              + 新建任务
            </button>
          </div>
        </div>
      </div>

      {/* 日志抽屉 */}
      {showLog && <SendLogDrawer logs={logs} sending={sending} onClose={() => setShowLog(false)} />}

      {/* 上传区 */}
      {tasks.length === 0 ? (
        <UploadZone
          dragging={dragging}
          onDrop={onDrop}
          onDragOver={onDragOver}
          onDragLeave={onDragLeave}
          onPickFile={() => fileRef.current?.click()}
          onDownloadTemplate={downloadTemplate}
          onNewTask={() => setShowNewTask(true)}
        />
      ) : (
        <>
          {/* 附加上传行 */}
          <div className="add-more-row">
            <button className="btn btn-ghost btn-sm" onClick={() => fileRef.current?.click()}>
              📎 继续导入表格
            </button>
            <button className="btn btn-ghost btn-sm" onClick={downloadTemplate}>
              ⬇ 下载模版
            </button>
          </div>

          {/* 任务表格 */}
          <div className="task-table-wrap card">
            <table className="task-table">
              <thead>
                <tr>
                  <th className="th-check">
                    <label className="checkbox-wrap">
                      <input
                        type="checkbox"
                        checked={allSelected}
                        ref={el => { if (el) el.indeterminate = someSelected && !allSelected }}
                        onClick={e => { e.preventDefault(); toggleAll() }}
                      />
                      <span className="checkmark" />
                    </label>
                  </th>
                  <th>联系人/群聊</th>
                  <th>消息类型</th>
                  <th>文字内容</th>
                  <th>发送时间</th>
                  <th>重复</th>
                  <th>状态</th>
                  <th className="th-action">操作</th>
                </tr>
              </thead>
              <tbody>
                {tasks.map(task => (
                  <TaskRow
                    key={task.id}
                    task={task}
                    checked={selected.has(task.id)}
                    onToggle={(currentlyChecked) => toggleOne(task.id, currentlyChecked)}
                    onEdit={() => { setEditTask(task); setShowNewTask(true) }}
                    onDelete={() => deleteTask(task.id)}
                    sending={sending}
                  />
                ))}
              </tbody>
            </table>
          </div>
        </>
      )}

      {/* 隐藏文件输入 */}
      <input
        ref={fileRef}
        type="file"
        accept=".xlsx,.xls"
        style={{ display: 'none' }}
        onChange={e => { handleFile(e.target.files[0]); e.target.value = '' }}
      />

      {/* 新建/编辑弹窗 */}
      {showNewTask && (
        <NewTaskModal
          initial={editTask}
          onConfirm={handleNewTask}
          onClose={() => { setShowNewTask(false); setEditTask(null) }}
        />
      )}
    </div>
  )
}

// ── 上传区 ────────────────────────────────────────────────
function UploadZone({ dragging, onDrop, onDragOver, onDragLeave, onPickFile, onDownloadTemplate, onNewTask }) {
  return (
    <div className="upload-zone-wrapper">
      <div
        className={`upload-zone ${dragging ? 'dragging' : ''}`}
        onDrop={onDrop}
        onDragOver={onDragOver}
        onDragLeave={onDragLeave}
        onClick={onPickFile}
      >
        <div className="upload-icon">
          <svg width="48" height="48" viewBox="0 0 24 24" fill="none">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8l-6-6z" fill="#07C160" opacity=".3"/>
            <path d="M14 2v6h6" stroke="#07C160" strokeWidth="1.5" strokeLinecap="round"/>
            <path d="M12 11v6M9 14l3-3 3 3" stroke="#07C160" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
          </svg>
        </div>
        <p className="upload-title">拖拽 Excel 表格到此处，或<span className="upload-link">点击选择文件</span></p>
        <p className="upload-sub">支持 .xlsx / .xls 格式</p>
      </div>

      <div className="upload-actions">
        <button className="btn btn-ghost" onClick={e => { e.stopPropagation(); onDownloadTemplate() }}>
          ⬇ 下载表格模版
        </button>
        <span className="upload-or">或者</span>
        <button className="btn btn-primary" onClick={e => { e.stopPropagation(); onNewTask() }}>
          + 手动新建任务
        </button>
      </div>
    </div>
  )
}

// ── 任务行 ────────────────────────────────────────────────
function TaskRow({ task, checked, onToggle, onEdit, onDelete, sending }) {
  const isSending = task.status === 'sending'
  const isSuccess = task.status === 'success'
  const canCheck = !isSuccess && !isSending

  return (
    <tr className={`task-row status-${task.status} ${checked ? 'row-checked' : ''}`}>
      <td className="td-check">
        <label className="checkbox-wrap" onClick={e => e.stopPropagation()}>
          <input
            type="checkbox"
            checked={checked}
            onChange={e => { e.stopPropagation(); onToggle(e.target.checked) }}
            disabled={!canCheck}
          />
          <span className="checkmark" />
        </label>
      </td>
      <td className="td-target">
        <span className="target-name">{task.target || '—'}</span>
      </td>
      <td>
        <MsgTypeBadge type={task.msg_type} />
      </td>
      <td className="td-text">
        <span title={task.text}>{task.text ? task.text.slice(0, 24) + (task.text.length > 24 ? '…' : '') : <em>（无文字）</em>}</span>
      </td>
      <td className="td-time">
        {task.send_time || <span className="text-dim">立即</span>}
      </td>
      <td>
        {task.repeat ? <span className="repeat-badge">{task.repeat}</span> : <span className="text-dim">—</span>}
      </td>
      <td>
        <StatusBadge status={task.status} />
      </td>
      <td className="td-action">
        <div className="action-btns">
          {!sending && !isSuccess && (
            <button className="icon-btn" title="编辑" onClick={onEdit}>
              <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor">
                <path d="M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zM20.71 7.04c.39-.39.39-1.02 0-1.41l-2.34-2.34c-.39-.39-1.02-.39-1.41 0l-1.83 1.83 3.75 3.75 1.83-1.83z"/>
              </svg>
            </button>
          )}
          {!sending && (
            <button className="icon-btn icon-btn-danger" title="删除" onClick={onDelete}>
              <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor">
                <path d="M6 19c0 1.1.9 2 2 2h8c1.1 0 2-.9 2-2V7H6v12zM19 4h-3.5l-1-1h-5l-1 1H5v2h14V4z"/>
              </svg>
            </button>
          )}
          {isSending && <div className="row-spinner" />}
        </div>
      </td>
    </tr>
  )
}

function MsgTypeBadge({ type }) {
  const styles = {
    '文字':    { bg: '#E6F7FF', color: '#0891B2' },
    '图片':    { bg: '#F3E8FF', color: '#7C3AED' },
    '文字+图片': { bg: '#FFF0F6', color: '#DB2777' },
  }
  const s = styles[type] || { bg: '#F5F5F5', color: '#888' }
  return <span className="badge" style={{ background: s.bg, color: s.color }}>{type || '—'}</span>
}

function StatusBadge({ status }) {
  const map = {
    waiting: ['badge-waiting', '待发送'],
    sending: ['badge-running', '发送中...'],
    success: ['badge-success', '✓ 成功'],
    failed:  ['badge-failed', '✗ 失败'],
    stopped: ['badge-stopped', '已停止'],
  }
  const [cls, label] = map[status] || ['', status]
  return <span className={`badge ${cls}`}>{label}</span>
}
