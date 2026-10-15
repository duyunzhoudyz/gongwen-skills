<#
gongwen-skills 一键安装脚本（Windows PowerShell 5.1 及以上）

用法：在仓库根目录打开 PowerShell，运行
  powershell -ExecutionPolicy Bypass -File .\install.ps1                    # 安装到本机探测到的全部 Agent 工具
  powershell -ExecutionPolicy Bypass -File .\install.ps1 -WhatIf            # 只演练，不写入任何文件
  powershell -ExecutionPolicy Bypass -File .\install.ps1 -InstallDeps       # 同时安装排版依赖
  powershell -ExecutionPolicy Bypass -File .\install.ps1 --pack hospital    # 加装行业包（可选 hospital、university、k12、soe、grassroots、all）
  powershell -ExecutionPolicy Bypass -File .\install.ps1 --target qwen      # 只装到指定工具
  powershell -ExecutionPolicy Bypass -File .\install.ps1 --uninstall        # 列出可卸载的 skill，加 --yes 才删除

本脚本做三件事：
  1. 检查 Python 3.9 及以上，没有就给出下载指引；
  2. 检查排版依赖 python-docx、openpyxl，只有你同意或传了 -InstallDeps 时才执行 pip install；
  3. 调用 python tools\install.py --target auto（你自己传了 --target 时以你的为准）。
除 -InstallDeps、-WhatIf 外，其余参数原样传给 tools\install.py，可用 --help 查看全部参数。
本脚本自身不联网；只有执行 pip install 时，pip 会联网下载依赖。
#>
param(
    [switch]$InstallDeps,
    [switch]$WhatIf
)

$ErrorActionPreference = 'Continue'
$RepoRoot = $PSScriptRoot
if (-not $RepoRoot) { $RepoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path }
$InstallPy = Join-Path $RepoRoot 'tools\install.py'
$DepsPackages = @('python-docx', 'openpyxl')
$PipMirror = 'https://pypi.tuna.tsinghua.edu.cn/simple'

function Write-Step([string]$Text) {
    Write-Host ''
    Write-Host "== $Text ==" -ForegroundColor Cyan
}

function Write-Warn([string]$Text) {
    Write-Host $Text -ForegroundColor Yellow
}

# 判断能否向用户提问：非交互会话、-NonInteractive 启动、输入被重定向时都不提问
function Test-Interactive {
    if (-not [Environment]::UserInteractive) { return $false }
    foreach ($a in [Environment]::GetCommandLineArgs()) {
        if ($a -like '-NonI*') { return $false }
    }
    try { if ([Console]::IsInputRedirected) { return $false } } catch { }
    return $true
}

# 依次尝试 python、py -3、python3，返回第一个 3.9 及以上的解释器
function Find-Python {
    $candidates = @(
        @{ Exe = 'python'; Pre = @() },
        @{ Exe = 'py'; Pre = @('-3') },
        @{ Exe = 'python3'; Pre = @() }
    )
    $script:TooOldPythons = @()
    foreach ($c in $candidates) {
        if (-not (Get-Command $c.Exe -ErrorAction SilentlyContinue)) { continue }
        $pre = $c.Pre
        $out = $null
        $code = 1
        try {
            # 未安装 Python 时，python 可能是微软商店的占位程序：带参数运行只报错退出，不会打开商店
            $out = & $c.Exe @pre -c "import sys; v=sys.version_info; print('%d.%d.%d' % (v[0], v[1], v[2])); sys.exit(0 if v >= (3, 9) else 1)" 2>$null
            $code = $LASTEXITCODE
        } catch {
            continue
        }
        if ($null -eq $out) { continue }
        $ver = ([string]($out | Select-Object -Last 1)).Trim()
        if ($ver -notmatch '^\d+\.\d+') { continue }
        if ($code -eq 0) {
            return @{ Exe = $c.Exe; Pre = $pre; Version = $ver }
        }
        $script:TooOldPythons += "$($c.Exe) $ver"
    }
    return $null
}

function Show-PythonGuide {
    Write-Warn '未找到 Python 3.9 及以上版本。'
    if ($script:TooOldPythons.Count -gt 0) {
        Write-Warn ("找到的版本过旧：" + ($script:TooOldPythons -join '、'))
    }
    Write-Host '安装方法：'
    Write-Host '  1. 打开 https://www.python.org/downloads/windows/ ，下载 3.9 及以上版本的安装包（推荐 3.12）；'
    Write-Host '  2. 安装第一屏务必勾选 "Add python.exe to PATH"，再点 Install Now；'
    Write-Host '  3. 装好后关闭并重新打开 PowerShell，再运行本脚本。'
    Write-Host '  也可以用系统自带的 winget 安装：winget install Python.Python.3.12'
    Write-Host '暂时不装 Python 也能用：'
    Write-Host '  Claude Code 用户：在 Claude Code 中运行 /plugin marketplace add duyunzhoudyz/gongwen-skills'
    Write-Host '  网页大模型用户：打开 dist\prompts\，把提示词文件内容复制到对话开头。'
}

# 在 PowerShell 内调用时，--pack hospital,soe 会被解析成数组，这里还原为逗号分隔的字符串
$PassArgs = @()
foreach ($a in $args) {
    if ($a -is [array]) { $PassArgs += ($a -join ',') } else { $PassArgs += [string]$a }
}
$DryRun = $WhatIf.IsPresent -or ($PassArgs -contains '--dry-run')
$HasTarget = @($PassArgs | Where-Object { $_ -eq '--target' -or $_ -like '--target=*' }).Count -gt 0
$ShowHelp = @($PassArgs | Where-Object { $_ -eq '-h' -or $_ -eq '--help' }).Count -gt 0
# 只查询、卸载或看帮助时，不需要检查排版依赖
$SkipDeps = $ShowHelp -or @($PassArgs | Where-Object { $_ -eq '--detect' -or $_ -eq '--list-targets' -or $_ -eq '--uninstall' }).Count -gt 0

# 让 Python 的中文输出在控制台和重定向文件中都不乱码；脚本结束时恢复原设置
$OldOutputEncoding = $null
try {
    $OldOutputEncoding = [Console]::OutputEncoding
    [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false
} catch { }
$OldPythonIoEncoding = $env:PYTHONIOENCODING
$env:PYTHONIOENCODING = 'utf-8'

$ExitCode = 0
try {
    # do { } while ($false) 只执行一次，遇到错误时用 break 提前结束，保证 finally 能恢复控制台设置
    do {
        if ($ShowHelp) {
            Write-Host 'install.ps1 自身参数：-WhatIf 只演练不写入；-InstallDeps 自动安装排版依赖。以下为 tools\install.py 的参数：'
        }
        if ($DryRun) {
            Write-Host '[演练模式] 只打印将要执行的操作，不安装依赖、不写入任何文件。' -ForegroundColor Cyan
        }

        if (-not (Test-Path -LiteralPath $InstallPy)) {
            Write-Warn "找不到 $InstallPy。请在 gongwen-skills 仓库根目录运行本脚本，并确认下载的文件完整。"
            $ExitCode = 2
            break
        }

        Write-Step '检查 Python'
        $py = Find-Python
        if ($null -eq $py) {
            Show-PythonGuide
            $ExitCode = 1
            break
        }
        $pyPre = @($py.Pre)
        Write-Host ("已找到 Python {0}（{1}）" -f $py.Version, ((@($py.Exe) + $pyPre) -join ' '))

        if (-not $SkipDeps) {
            Write-Step '检查排版依赖（python-docx、openpyxl）'
            $probe = "import importlib.util as u; print(','.join(n for n, m in (('python-docx', 'docx'), ('openpyxl', 'openpyxl')) if u.find_spec(m) is None))"
            $missingText = ''
            try {
                $missingOut = & $py.Exe @pyPre -c $probe 2>$null
                if ($null -ne $missingOut) { $missingText = ([string]($missingOut | Select-Object -Last 1)).Trim() }
            } catch {
                $missingText = ($DepsPackages -join ',')
            }
            if ($missingText -eq '') {
                Write-Host '依赖齐全，排版功能可以直接使用。'
            } else {
                $pipShown = (@($py.Exe) + $pyPre + @('-m', 'pip', 'install') + $DepsPackages) -join ' '
                Write-Warn ("缺少：" + ($missingText -replace ',', '、') + "。不装也能写公文，只是排版会改为输出 WPS/Word 手动设置清单。")
                $doInstall = $false
                if ($DryRun) {
                    Write-Host "[演练] 不安装依赖。正式安装时可加 -InstallDeps，或手动运行：$pipShown"
                } elseif ($InstallDeps) {
                    $doInstall = $true
                } elseif (Test-Interactive) {
                    $answer = Read-Host '是否现在用 pip 安装？输入 y 确认，直接回车跳过'
                    $doInstall = ($answer -match '^\s*(y|yes|是)\s*$')
                    if (-not $doInstall) { Write-Host "已跳过。需要时运行：$pipShown" }
                } else {
                    Write-Host "已跳过。需要时运行：$pipShown（或给本脚本加 -InstallDeps）"
                }
                if ($doInstall) {
                    Write-Host "正在运行：$pipShown"
                    $pipArgs = @('-m', 'pip', 'install') + $DepsPackages
                    & $py.Exe @pyPre @pipArgs
                    if ($LASTEXITCODE -ne 0) {
                        Write-Warn '依赖安装失败，不影响安装 skill。可以稍后手动处理：'
                        Write-Host "  网络较慢时改用国内镜像：$pipShown -i $PipMirror"
                        Write-Host ("  提示没有 pip 时，先运行：" + ((@($py.Exe) + $pyPre + @('-m', 'ensurepip', '--upgrade')) -join ' '))
                    } else {
                        Write-Host '排版依赖安装完成。'
                    }
                }
            }
        }

        if ($PassArgs -contains '--uninstall') { Write-Step '卸载 skill' } else { Write-Step '安装 skill' }
        $finalArgs = @()
        if (-not $HasTarget) { $finalArgs += @('--target', 'auto') }
        $finalArgs += $PassArgs
        if ($WhatIf.IsPresent -and -not ($PassArgs -contains '--dry-run')) { $finalArgs += '--dry-run' }
        & $py.Exe @pyPre $InstallPy @finalArgs
        $ExitCode = $LASTEXITCODE
        if ($ExitCode -eq 3) {
            Write-Warn '可以先安装 Claude Code、Qwen Code 等 Agent 工具后重新运行，或用 --target 指定目标。'
        } elseif ($ExitCode -ne 0) {
            Write-Warn "安装器返回错误（退出码 $ExitCode），请查看上方提示。"
        }
    } while ($false)
} finally {
    if ($null -ne $OldOutputEncoding) {
        try { [Console]::OutputEncoding = $OldOutputEncoding } catch { }
    }
    $env:PYTHONIOENCODING = $OldPythonIoEncoding
}
exit $ExitCode
