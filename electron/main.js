const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron')
const path = require('path')
const { spawn, spawnSync } = require('child_process')
const fs = require('fs')
const os = require('os')

// ─── Python 路径（运行时检测，避免 isPackaged 在模块加载时出错）───────────
function getPythonPath() {
  if (!app.isPackaged) {
    return process.platform === 'win32' ? 'python' : 'python3'
  }
  // extraResources: python/ 映射到 Resources/python/
  const base = path.join(process.resourcesPath, 'python')
  if (process.platform === 'win32') {
    return path.join(base, 'python.exe')
  }
  return path.join(base, 'bin', 'python3')
}

function getScriptPath(name) {
  // extraResources: python/ -> Resources/python/
  if (!app.isPackaged) {
    return path.join(__dirname, '..', 'python', name)
  }
  return path.join(process.resourcesPath, 'python', name)
}

// ─── 加载 dist/（生产）vs localhost:5173（开发）────────────────────────
function loadMainWindow(win) {
  const distIndex = path.join(__dirname, '..', 'dist', 'index.html')
  if (fs.existsSync(distIndex)) {
    win.loadFile(distIndex)
  } else {
    win.loadURL('http://localhost:5173')
  }
}

// ─── 主进程 ───────────────────────────────────────────────────────────
let daemonProcess = null

function createWindow() {
  const win = new BrowserWindow({
    width: 980, height: 680,
    minWidth: 800, minHeight: 560,
    titleBarStyle: process.platform === 'darwin' ? 'hiddenInset' : 'default',
    backgroundColor: '#F5F5F5',
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js'),
    },
    icon: path.join(__dirname, '..', 'assets', 'icon.png'),
  })
  loadMainWindow(win)
  return win
}

// ─── IPC ──────────────────────────────────────────────────────────────

ipcMain.handle('get-config', async () => {
  try {
    const p = path.join(os.homedir(), '.wechat-sender', 'config.json')
    if (fs.existsSync(p)) return JSON.parse(fs.readFileSync(p, 'utf8'))
  } catch (_) {}
  return { send_interval: 5, max_per_minute: 8, poll_seconds: 15, dry_run: false }
})

ipcMain.handle('save-config', async (_, cfg) => {
  try {
    const dir = path.join(os.homedir(), '.wechat-sender')
    fs.mkdirSync(dir, { recursive: true })
    fs.writeFileSync(path.join(dir, 'config.json'), JSON.stringify(cfg, null, 2), 'utf8')
    return { ok: true }
  } catch (e) {
    return { ok: false, error: e.message }
  }
})

ipcMain.handle('parse-excel', async (_, bytes) => {
  const tmp = path.join(os.tmpdir(), 'wst_upload_' + Date.now() + '.xlsx')
  try {
    fs.writeFileSync(tmp, Buffer.from(bytes))
    const pythonBin = getPythonPath()
    const scriptPath = getScriptPath('app/parse_excel.py')
    console.log('[parse-excel] python:', pythonBin, 'script:', scriptPath, 'tmp:', tmp)
    const r = spawnSync(pythonBin, [scriptPath, tmp], {
      encoding: 'utf8', timeout: 15000,
      env: Object.assign({}, process.env, { PYTHONIOENCODING: 'utf-8' }),
    })
    fs.unlinkSync(tmp)
    if (r.status !== 0) {
      console.error('[parse-excel] python error:', r.stderr)
      return { error: r.stderr }
    }
    return JSON.parse(r.stdout)
  } catch (e) {
    console.error('[parse-excel] exception:', e)
    return { error: e.message }
  }
})

ipcMain.handle('download-template', async (event) => {
  const win = BrowserWindow.fromWebContents(event.sender)
  const result = await dialog.showSaveDialog(win, {
    title: '保存表格模版',
    defaultPath: path.join(os.homedir(), 'Desktop', 'wechat_template.xlsx'),
    filters: [{ name: 'Excel', extensions: ['xlsx'] }],
  })
  if (result.canceled) return
  const r = spawnSync(getPythonPath(), [getScriptPath('app/template_gen.py'), result.filePath], {
    encoding: 'utf8',
    env: Object.assign({}, process.env, { PYTHONIOENCODING: 'utf-8' }),
  })
  return r.status === 0
    ? { ok: true, path: result.filePath }
    : { ok: false, err: r.stderr }
})

ipcMain.handle('send-now', async (event) => {
  const win = BrowserWindow.fromWebContents(event.sender)
  const proc = spawn(getPythonPath(), [getScriptPath('app/cli.py'), 'send'], {
    env: Object.assign({}, process.env, { PYTHONIOENCODING: 'utf-8', PYTHONUNBUFFERED: '1' }),
    windowsHide: true,
  })
  proc.stdout.on('data', function(d) {
    win.webContents.send('send-progress', d.toString('utf8'))
  })
  proc.stderr.on('data', function(d) {
    win.webContents.send('send-progress', '[ERROR] ' + d.toString('utf8'))
  })
  proc.on('close', function(code) {
    win.webContents.send('send-done', { code: code })
  })
  return { ok: true }
})

ipcMain.handle('send-selected', async (event, taskData) => {
  const win = BrowserWindow.fromWebContents(event.sender)
  const jsonArg = Array.isArray(taskData) ? JSON.stringify(taskData) : '[]'
  const pythonBin = getPythonPath()
  const scriptPath = getScriptPath('app/cli.py')
  console.log('[send-selected] python:', pythonBin)
  console.log('[send-selected] script:', scriptPath)
  console.log('[send-selected] tasks:', jsonArg.slice(0, 80))

  const { execFileSync } = require('child_process')
  try {
    const stdout = execFileSync(pythonBin, [scriptPath, 'send', '--tasks-json', jsonArg], {
      encoding: 'utf8',
      timeout: 0,
      env: Object.assign({}, process.env, { PYTHONIOENCODING: 'utf-8' }),
      windowsHide: true,
    })
    console.log('[send-selected] stdout length:', stdout.length)
    // 发送所有输出到日志
    win.webContents.send('send-progress', stdout)
    // 逐行解析任务结果
    stdout.split('\n').forEach(function(line) {
      if (!line.trim()) return
      const hasSuccess = line.includes('✅')
      const hasFailed = line.includes('❌')
      if (hasSuccess || hasFailed) {
        const taskMatch = line.match(/^\[(\d+)\/(\d+)\] → (.+?) \[(.+?)\]/)
        if (taskMatch) {
          win.webContents.send('task-status-update', {
            target: taskMatch[3].trim(),
            msg_type: taskMatch[4].trim(),
            status: hasSuccess ? 'success' : 'failed',
          })
        }
      }
    })
    win.webContents.send('send-done', { code: 0 })
  } catch(e) {
    console.error('[send-selected] exception:', e.message)
    const stderr = e.stderr ? e.stderr.toString() : ''
    const stdout = e.stdout ? e.stdout.toString() : ''
    if (stdout) win.webContents.send('send-progress', stdout)
    if (stderr) win.webContents.send('send-progress', '[ERROR] ' + stderr)
    win.webContents.send('send-done', { code: e.status || 1, error: e.message })
  }
  return { ok: true }
})

ipcMain.handle('stop-send', async () => {
  try {
    fs.writeFileSync(path.join(os.homedir(), '.wechat-sender', 'stop_signal'), String(Date.now()))
  } catch (_) {}
  return { ok: true }
})

ipcMain.handle('daemon-start', async (event) => {
  if (daemonProcess) return { ok: false, error: '已在运行' }
  const win = BrowserWindow.fromWebContents(event.sender)
  daemonProcess = spawn(getPythonPath(), [getScriptPath('app/cli.py'), 'daemon'], {
    env: Object.assign({}, process.env, { PYTHONIOENCODING: 'utf-8', PYTHONUNBUFFERED: '1' }),
    windowsHide: true,
  })
  daemonProcess.stdout.on('data', function(d) {
    win.webContents.send('daemon-log', d.toString('utf8'))
  })
  daemonProcess.stderr.on('data', function(d) {
    win.webContents.send('daemon-log', '[ERR] ' + d.toString('utf8'))
  })
  daemonProcess.on('close', function() {
    daemonProcess = null
    win.webContents.send('daemon-stopped')
  })
  return { ok: true }
})

ipcMain.handle('daemon-stop', async () => {
  if (daemonProcess) { daemonProcess.kill(); daemonProcess = null }
  return { ok: true }
})

ipcMain.handle('daemon-status', async () => ({ running: daemonProcess !== null }))

ipcMain.handle('add-schedule', async (_, schedule) => {
  const f = path.join(os.homedir(), '.wechat-sender', 'schedules.json')
  let arr = []
  try { arr = JSON.parse(fs.readFileSync(f, 'utf8')) } catch (_) {}
  arr.push(schedule)
  fs.writeFileSync(f, JSON.stringify(arr, null, 2))
  return { ok: true }
})

ipcMain.handle('remove-schedule', async (_, id) => {
  const f = path.join(os.homedir(), '.wechat-sender', 'schedules.json')
  let arr = []
  try { arr = JSON.parse(fs.readFileSync(f, 'utf8')) } catch (_) {}
  arr = arr.filter(function(s) { return s.id !== id })
  fs.writeFileSync(f, JSON.stringify(arr, null, 2))
  return { ok: true }
})

ipcMain.handle('pick-file', async (event) => {
  const win = BrowserWindow.fromWebContents(event.sender)
  const r = await dialog.showOpenDialog(win, {
    filters: [{ name: 'Excel', extensions: ['xlsx', 'xls'] }],
    properties: ['openFile'],
  })
  return r.canceled ? null : r.filePaths[0]
})

ipcMain.handle('open-file', async (_, filePath) => {
  await shell.openPath(filePath)
})

ipcMain.handle('get-platform', () => process.platform)

// ─── App lifecycle ────────────────────────────────────────────────────
app.whenReady().then(function() {
  createWindow()
  app.on('activate', function() {
    if (BrowserWindow.getAllWindows().length === 0) createWindow()
  })
})

app.on('window-all-closed', function() {
  if (daemonProcess) daemonProcess.kill()
  if (process.platform !== 'darwin') app.quit()
})

app.on('before-quit', function() {
  if (daemonProcess) daemonProcess.kill()
})
