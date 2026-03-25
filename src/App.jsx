import { useState, useEffect } from 'react'
import Sidebar from './components/Sidebar'
import TitleBar from './components/TitleBar'
import TasksPage from './pages/TasksPage'
import SettingsPage from './pages/SettingsPage'
import TemplatePage from './pages/TemplatePage'
import './styles/app.css'

const PAGES = {
  tasks: TasksPage,
  settings: SettingsPage,
  template: TemplatePage,
}

export default function App() {
  const [page, setPage] = useState('tasks')
  const PageComp = PAGES[page] || TasksPage

  // 启动时自动拉起守护进程
  useEffect(() => {
    if (window.api?.daemonStart) {
      window.api.daemonStart().catch(() => {})
    }
  }, [])

  return (
    <div className="app-shell">
      <TitleBar />
      <div className="app-body">
        <Sidebar current={page} onNavigate={setPage} />
        <main className="app-content">
          <PageComp />
        </main>
      </div>
    </div>
  )
}
