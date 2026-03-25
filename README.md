# 微信发送助手 WeChat Sender

跨平台 Electron 桌面工具，支持从 Excel 批量导入发送任务、手动创建任务、Daemon 守护模式、定时发送。

![Platform](https://img.shields.io/badge/platform-macOS%20%7C%20Windows-blue)
![License](https://img.shields.io/badge/license-MIT-green)

## 功能特性

- **Excel 批量导入** — 下载标准模板，填写后一键导入任务列表
- **手动创建任务** — 直接在应用内添加单个发送目标
- **多选发送** — 勾选任务后批量发送，支持按目标+类型匹配
- **Daemon 守护模式** — 后台常驻，自动检测定时任务并执行
- **实时发送日志** — 每条任务的发送进度和结果实时显示
- **任务状态同步** — 发送完成后自动更新任务列表状态
- **微信原生自动化** — macOS AppleScript / Windows Uiautomation

## 系统要求

- macOS 10.15+ 或 Windows 10+
- 已安装微信（用于自动化发送）

## 安装包下载

Releases 页面：https://github.com/howtimeschange/wechat-sender/releases

## 开发

```bash
# macOS
npm install
npm run dev           # 开发热重载模式
npm run dist:mac      # 构建 macOS DMG 安装包

# Windows
npm install
npm run dist:win      # 构建 Windows NSIS 安装包
```

## Excel 模板说明

| 字段 | 说明 |
|------|------|
| * 应用 | 微信 / 钉钉 / 飞书 |
| * 联系人/群聊 | 精确的联系人昵称或群聊名称 |
| * 消息类型 | 文字 / 图片 / 文字+图片 |
| * 文字内容 | 支持变量：{name} {date} {time} |
| 图片路径 | 本机绝对路径（消息类型含图片时必填） |
| 发送时间 | 留空=立即发送；格式 YYYY-MM-DD HH:MM |
| 重复 | 留空=单次；daily=每天；weekly=每周；workday=工作日 |

## 项目结构

```
wechat-sender-gui/
├── electron/          # Electron 主进程
│   ├── main.js         # IPC handlers、窗口管理
│   └── preload.js      # 上下文桥接 API
├── python/
│   └── app/
│       ├── cli.py      # 核心 CLI（发送逻辑、AppleScript）
│       ├── parse_excel.py    # Excel 解析
│       └── template_gen.py   # 模板生成
├── src/
│   ├── pages/          # React 页面
│   └── components/      # UI 组件
└── assets/             # 图标等资源
```

## 许可证

MIT
