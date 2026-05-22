' 静默启动元数据清除工具（无终端窗口）
Set sh = CreateObject("WScript.Shell")
Set fs = CreateObject("Scripting.FileSystemObject")
dir = fs.GetParentFolderName(WScript.ScriptFullName)
sh.Run "pythonw """ & dir & "\metadata_cleaner.py""", 0, False
