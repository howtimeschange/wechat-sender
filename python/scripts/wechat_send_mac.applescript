on run argv
	if (count of argv) < 4 then
		error "Usage: targetName msgType msgText imagePath"
	end if

	set targetName to item 1 of argv
	set msgType to item 2 of argv -- 文字 | 图片 | 文字+图片
	set msgText to item 3 of argv
	set imagePath to item 4 of argv

	try
		tell application "WeChat" to activate
	on error
		tell application id "com.tencent.xinWeChat" to activate
	end try
	delay 0.8

	tell application "System Events"
		set pName to ""
		if exists process "WeChat" then
			set pName to "WeChat"
		else if exists process "微信 3" then
			set pName to "微信 3"
		else
			error "未找到微信进程（请确认微信已启动）"
		end if

		tell process pName
			set frontmost to true
			delay 0.3

			-- 打开搜索框
			keystroke "f" using {command down}
			delay 0.25

			-- 粘贴目标名称
			set the clipboard to targetName
			keystroke "v" using {command down}
			delay 0.25

			-- 回车进入目标会话
			key code 36
			delay 0.7

			-- ============================================================
			-- 【FIX】若当前窗口已经是目标会话（发送后聊天框出现重复消息）
			-- 先按 Escape 关闭可能的搜索残留浮层
			-- ============================================================
			key code 53
			delay 0.15

			-- ============================================================
			-- 文字发送
			-- ============================================================
			if msgType is "文字" or msgType is "文字+图片" then
				if msgText is not "" then
					set the clipboard to msgText
					delay 0.05
					keystroke "v" using {command down}
					delay 0.15
					key code 36
					delay 0.35
				end if
			end if

			-- ============================================================
			-- 【FIX】图片发送 — 剪贴板 alias 粘贴方案
			-- 绕过 Cmd+O 文件对话框在不同微信版本的行为不一致问题
			-- ============================================================
			if msgType is "图片" or msgType is "文字+图片" then
				if imagePath is not "" then
					try
						set imgAlias to (POSIX file imagePath) as alias
						set the clipboard to imgAlias
						delay 0.15
						keystroke "v" using {command down}
						delay 0.25
						key code 36
						delay 0.4
					on error
						error "图片路径无效或文件不存在: " & imagePath
					end try
				else
					error "消息类型包含图片，但图片路径为空"
				end if
			end if
		end tell
	end tell
end run
