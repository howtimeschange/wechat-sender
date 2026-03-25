on run argv
	if (count of argv) < 4 then
		error "Usage: targetName msgType msgText imagePath"
	end if

	set targetName to item 1 of argv
	set msgType to item 2 of argv -- 文字 | 图片 | 文字+图片
	set msgText to item 3 of argv
	set imagePath to item 4 of argv

	-- 激活微信
	try
		tell application "WeChat" to activate
	on error
		tell application id "com.tencent.xinWeChat" to activate
	end try
	delay 0.8

	-- 查找微信进程名称
	set pName to ""
	tell application "System Events"
		if exists process "WeChat" then
			set pName to "WeChat"
		else if exists process "微信 3" then
			set pName to "微信 3"
		else
			error "未找到微信进程（请确认微信已启动）"
		end if
	end tell

	-- 发送标志：任何一步失败都设为 false
	set sendSucceeded to true

	-- 打开搜索框并搜索联系人
	tell application "System Events"
		tell process pName
			set frontmost to true
			delay 0.3

			keystroke "f" using {command down}
			delay 0.25

			set the clipboard to targetName
			keystroke "v" using {command down}
			delay 0.25

			key code 36 -- 回车进入会话
			delay 0.7

			-- 关闭可能的搜索残留浮层
			key code 53
			delay 0.15

			-- 文字发送
			if msgType is "文字" or msgType is "文字+图片" then
				if msgText is not "" then
					set the clipboard to msgText
					delay 0.05
					keystroke "v" using {command down}
					delay 0.15
					key code 36
					delay 0.35
				else
					set sendSucceeded to false
				end if
			end if

			-- 图片发送
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
					on error errMsg
						set sendSucceeded to false
						error "图片发送失败: " & errMsg
					end try
				else
					set sendSucceeded to false
					error "消息类型包含图片，但图片路径为空"
				end if
			end if
		end tell
	end tell

	-- 最终检查：若有任何步骤标记失败，则以错误退出
	if sendSucceeded is false then
		error "微信发送未完成（内容为空或发送被中断）"
	end if
end run
