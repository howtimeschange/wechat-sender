# 微信发送助手 WeChat Sender

跨平台 Electron 桌面工具，支持从 Excel 批量导入发送任务、手动创建任务、Daemon 守护模式、定时发送。

![Platform](https://img.shields.io/badge/platform-macOS%20%7C%20Windows-blue)
![License](https://img.shields.io/badge/license-MIT-green)
![Version](https://img.shields.io/badge/version-v1.1.31-blue)

![WeChat Sender Demo](https://raw.githubusercontent.com/howtimeschange/wechat-sender/main/png/demo.png)

## 功能特性

- **Excel 批量导入** — 下载标准模板，填写后一键导入任务列表
- **手动创建任务** — 直接在应用内添加单个发送目标
- **多选发送** — 勾选任务后批量发送，支持按目标+类型匹配
- **Daemon 守护模式** — 后台常驻，自动检测定时任务并执行
- **实时发送日志** — 每条任务的发送进度和结果实时显示
- **任务状态同步** — 发送完成后自动更新任务列表状态
- **跨平台原生自动化** — macOS AppleScript / Windows pywinauto

## 系统要求

- macOS 10.15+（Apple Silicon Mac + Intel Mac 均支持）
- Windows 10/11（64 位，推荐 Windows Server 2022）
- 已安装并登录微信 PC 版（自动化操作依赖微信桌面端）

## 安装包下载

Releases 页面：https://github.com/howtimeschange/wechat-sender/releases

### 各版本说明

| 版本 | 说明 |
|------|------|
| `.dmg`（macOS） | 双击安装，自动识别 Apple Silicon / Intel |
| `.exe`（Windows NSIS） | 一键安装，安装完成后桌面快捷方式 |

## 技术架构

### macOS
- **自动化引擎**：AppleScript `osascript`，直接调起微信的 GUI 脚本接口
- **剪贴板**：`pyperclip` → `osascript`，零延迟
- **搜索**：`open location "weixin://search/?searchkey=..."` URL 协议，跳转无需模拟输入

### Windows
- **自动化引擎**：pywinauto（基于 Windows UI Automation API）
- **键盘输入**：`keyboard.send_keys()` → `SendMessage(WM_SETTEXT)`，直接写文本到控件，**不走剪贴板，不触发微信监控，无防抖问题**
- **剪贴板**：`pyperclip`，仅在发送图片时使用（图片只能通过剪贴板发送）
- **窗口识别**：枚举桌面所有窗口，按进程名 + 类名 + 面积评分，选出最可能是微信主窗口的那个

> **为什么 Windows 不用 `uiautomation`？**  
> `uiautomation.SendKeys()` 底层用 `SendInput`（键盘硬件模拟），部分 Edit 控件上存在兼容性问题，且对某些微信版本会触发崩溃。pywinauto 用 `SendMessage` 直接向控件窗口发字符消息，绕过键盘模拟层，更稳定。

## 开发

```bash
# macOS
npm install
npm run dev           # 开发热重载模式
npm run dist:mac      # 构建 macOS DMG 安装包

# Windows
npm install
npm run dev           # Windows 开发调试
npm run dist:win      # 构建 Windows NSIS 安装包
```

### 依赖说明

| 平台 | Python 依赖 |
|------|------------|
| macOS | openpyxl, PyYAML, rich（AppleScript 负责 GUI 自动化） |
| Windows | openpyxl, pywinauto, psutil, pyperclip, pywin32, Pillow |

## Excel 模板说明

| 字段 | 说明 |
|------|------|
| `* 应用` | 微信 / 钉钉 / 飞书（目前仅支持微信） |
| `* 联系人/群聊` | 精确的联系人昵称或群聊名称 |
| `* 消息类型` | 文字 / 图片 / 文字+图片 |
| `* 文字内容` | 支持变量：`{name}` `{date}` `{time}` |
| `图片路径` | 本机绝对路径（消息类型含图片时必填） |
| `发送时间` | 留空=立即发送；格式 `YYYY-MM-DD HH:MM` |
| `重复` | 留空=单次；`daily`=每天；`weekly`=每周；`workday`=工作日 |

## 项目结构

```
wechat-sender/
├── electron/              # Electron 主进程
│   ├── main.js            # IPC handlers、窗口管理
│   └── preload.js         # 上下文桥接 API
├── python/
│   ├── app/
│   │   ├── cli.py         # 核心 CLI（跨平台统一入口）
│   │   ├── parse_excel.py # Excel 解析
│   │   └── template_gen.py # 模板生成
│   └── scripts/
│       ├── wechat_send_mac.applescript  # macOS AppleScript
│       └── wechat_send_win.py           # Windows pywinauto 自动化
├── src/                   # React 页面
├── scripts/
│   └── afterPack.js      # electron-builder hook（macOS 打包后复制 Python 依赖）
└── .github/workflows/
    └── build.yml         # GitHub Actions 自动构建 + Release 发布
```

## 工作流程

1. 用户在 Excel 中填写任务列表，导入应用
2. 应用读取任务，按定时时间和频率控制排队
3. macOS：通过 `osascript` 执行 AppleScript 命令  
   Windows：通过 `pywinauto` 执行 UI 自动化
4. 每条任务执行后，结果回写入 Excel 状态列

## 常见问题

**Q: Windows 版提示"找不到微信窗口"？**  
确保微信 PC 版已启动并登录，且窗口未被最小化到托盘。

**Q: Windows 版搜索"文件传输助手"失败？**  
在微信搜索框手动测试一下能否搜到，如果微信本身搜索异常（如防抖设置变化），可能需要调整 `search_contact` 中的等待延时。

**Q: macOS 版发送失败，提示 AppleScript 权限？**  
系统设置 → 隐私与安全性 → 自动化 → 允许「微信发送助手」控制「微信」。

**Q: 批量发送频率限制？**  
默认每分钟最多 8 条，两条间隔默认 5 秒。可在设置中调整。

## 许可证

MIT
