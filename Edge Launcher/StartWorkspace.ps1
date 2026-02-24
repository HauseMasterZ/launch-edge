Add-Type -AssemblyName System.Windows.Forms

$mutex = New-Object System.Threading.Mutex($false, "Global\WorkspaceLauncher")
if (-not $mutex.WaitOne(0)) { exit 0 }

try {
    $scriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
    $configPath = Join-Path $scriptDir "workspace_config.json"

    if (-not (Test-Path $configPath)) {
        @{
            CloseExistingEdge = $true
            Middle = @("https://mail.google.com/mail/u/0/")
            Left   = @("https://discord.com/channels/@me", "https://photos.google.com/")
            Right  = @("https://web.whatsapp.com/", "https://www.perplexity.ai/")
        } | ConvertTo-Json | Out-File $configPath -Encoding UTF8
    }

    $cfg        = Get-Content $configPath -Raw | ConvertFrom-Json
    $MiddleURLs = @($cfg.Middle)
    $LeftURLs   = @($cfg.Left)
    $RightURLs  = @($cfg.Right)

    if (-not ("WinAPIWS_v5" -as [type])) {
        Add-Type @"
        using System;
        using System.Collections.Generic;
        using System.Runtime.InteropServices;
        using System.Text;
        public class WinAPIWS_v5 {
            [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
            [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
            [DllImport("user32.dll")] public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int W, int H, bool b);
            [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
            [DllImport("user32.dll")] public static extern int GetWindowTextLength(IntPtr hWnd);
            [DllImport("user32.dll")] public static extern int GetClassName(IntPtr hWnd, StringBuilder cls, int n);
            [DllImport("user32.dll")] [return: MarshalAs(UnmanagedType.Bool)] public static extern bool EnumWindows(EnumWindowsProc f, IntPtr l);
            [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
            [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
            public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
            public static bool IsEdgeWindow(IntPtr hWnd) {
                if (!IsWindowVisible(hWnd)) return false;
                var cls = new StringBuilder(256);
                GetClassName(hWnd, cls, 256);
                return cls.ToString() == "Chrome_WidgetWin_1" && GetWindowTextLength(hWnd) > 0;
            }
            public static List<IntPtr> GetEdgeWindows() {
                var list = new List<IntPtr>();
                EnumWindows((hWnd, lParam) => { if (IsEdgeWindow(hWnd)) list.Add(hWnd); return true; }, IntPtr.Zero);
                return list;
            }
        }
"@
    }

    [WinAPIWS_v5]::SetProcessDPIAware() | Out-Null

    $edge = @(
        "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        "C:\Program Files\Microsoft\Edge\Application\msedge.exe"
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1

    $primary     = [System.Windows.Forms.Screen]::PrimaryScreen
    $all         = [System.Windows.Forms.Screen]::AllScreens
    $leftScreen  = $all | Where-Object { $_.Bounds.X -lt $primary.Bounds.X } | Sort-Object { $_.Bounds.X } | Select-Object -Last 1
    $rightScreen = $all | Where-Object { $_.Bounds.X -gt $primary.Bounds.X } | Sort-Object { $_.Bounds.X } | Select-Object -First 1

    if ([bool]$cfg.CloseExistingEdge) {
        $edgeProcesses = Get-Process msedge -ErrorAction SilentlyContinue
        if ($edgeProcesses) {
            $edgeProcesses | Stop-Process -Force
            $edgeProcesses | Wait-Process -Timeout 5 -ErrorAction SilentlyContinue 
        }
    }

    function Wait-NewEdgeWindow($known) {
        $deadline = (Get-Date).AddSeconds(20)
        while ((Get-Date) -lt $deadline) {
            Start-Sleep -Milliseconds 50
            $new = [WinAPIWS_v5]::GetEdgeWindows() | Where-Object { $known -notcontains $_ }
            if ($new.Count -gt 0) { return $new[0] }
        }
        return [IntPtr]::Zero
    }

    $known = [WinAPIWS_v5]::GetEdgeWindows()
    $args  = "--new-window --disable-session-crashed-bubble --restore-last-session=false "

    Start-Process $edge -ArgumentList ($args + ($MiddleURLs -join " "))
    $hwnd_middle = Wait-NewEdgeWindow $known
    $known = [WinAPIWS_v5]::GetEdgeWindows()

    if ($LeftURLs.Count -gt 0 -and $leftScreen) {
        Start-Process $edge -ArgumentList ($args + ($LeftURLs -join " "))
        $hwnd_left = Wait-NewEdgeWindow $known
        $known = [WinAPIWS_v5]::GetEdgeWindows()
    }

    if ($RightURLs.Count -gt 0 -and $rightScreen) {
        Start-Process $edge -ArgumentList ($args + ($RightURLs -join " "))
        $hwnd_right = Wait-NewEdgeWindow $known
    }

    foreach ($a in @(
        @{ hwnd = $hwnd_middle; screen = $primary     },
        @{ hwnd = $hwnd_left;   screen = $leftScreen  },
        @{ hwnd = $hwnd_right;  screen = $rightScreen }
    )) {
        if (-not $a.hwnd -or $a.hwnd -eq [IntPtr]::Zero -or $null -eq $a.screen) { continue }
        [WinAPIWS_v5]::ShowWindow($a.hwnd, 9)
        [WinAPIWS_v5]::MoveWindow($a.hwnd, $a.screen.Bounds.X, $a.screen.Bounds.Y, $a.screen.Bounds.Width, $a.screen.Bounds.Height, $true)
        [WinAPIWS_v5]::ShowWindow($a.hwnd, 3)
    }

    if ($hwnd_right -and $hwnd_right -ne [IntPtr]::Zero) {
        [WinAPIWS_v5]::SetForegroundWindow($hwnd_right)
        $attempts = 0
        while ([WinAPIWS_v5]::GetForegroundWindow() -ne $hwnd_right -and $attempts -lt 10) {
            Start-Sleep -Milliseconds 50
            $attempts++
        }
        Start-Sleep -Milliseconds 100
        [System.Windows.Forms.SendKeys]::SendWait("^+,")
    }
}
finally {
    if ($mutex) {
        $mutex.ReleaseMutex()
        $mutex.Dispose()
    }
}