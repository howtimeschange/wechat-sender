/**
 * afterPack.js — electron-builder hook
 * 在打包完成后，将系统 Python 的 openpyxl / PyYAML / rich 复制到
 * app/Contents/Resources/python/lib/，让打包后的 Python 能 import 它们。
 */
const { execSync } = require('child_process')
const path = require('path')
const fs = require('fs')

module.exports = function afterPack(context) {
  const { appOutDir, packager } = context
  const platform = packager.platform.name

  if (platform !== 'Mac OS X') return

  // 找到系统 Python 的 site-packages
  let sitePackages
  try {
    sitePackages = execSync(
      'python3 -c "import site; print(site.getusersitepackages() or site.getsitepackages()[0])"',
      { encoding: 'utf8' }
    ).trim()
  } catch {
    // fallback: try main system site-packages
    try {
      sitePackages = execSync(
        'python3 -c "import sys; print(sys.path[1])"',
        { encoding: 'utf8' }
      ).trim()
    } catch {
      console.warn('[afterPack] Cannot detect Python site-packages, skipping openpyxl copy')
      return
    }
  }

  console.log('[afterPack] Detected site-packages:', sitePackages)

  // 目标 lib 目录
  const destLib = path.join(
    appOutDir,
    'WechatSender.app',
    'Contents',
    'Resources',
    'python',
    'lib',
    'python3.14',
    'site-packages'
  )

  fs.mkdirSync(destLib, { recursive: true })

  // 要打包的包（只取 openpyxl + 其依赖）
  const pkgs = ['openpyxl', 'PyYAML', 'rich', 'et_xmlfile', 'jinja2', 'markupsafe']

  for (const pkg of pkgs) {
    const src = path.join(sitePackages, pkg)
    const dst = path.join(destLib, pkg)
    if (fs.existsSync(src)) {
      console.log(`[afterPack] Copying ${pkg}...`)
      fs.rmSync(dst, { recursive: true, force: true })
      fs.cpSync(src, dst, { recursive: true })
    }
  }

  // 复制 zip_safe marker 文件等
  for (const fname of ['easy_install.py', 'site.py']) {
    const src = path.join(sitePackages, fname)
    if (fs.existsSync(src)) {
      fs.copyFileSync(src, path.join(destLib, fname))
    }
  }

  console.log('[afterPack] Done — openpyxl & deps copied to bundle.')
}
