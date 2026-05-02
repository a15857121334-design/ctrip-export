## Shell rule

在 Windows 上运行 PowerShell 命令时，默认使用 PowerShell 7：

```powershell
pwsh -NoProfile -Command "<command>"
```

不要直接使用 powershell.exe / Windows PowerShell 5.1，除非 pwsh 执行失败。
如果命令涉及中文输出，优先使用 pwsh 以减少编码问题。
