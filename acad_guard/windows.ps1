# Windows of AutoCAD (and windows owned by it / crash reporters / Windows "open with"):
# list them with their UI elements, press a button, close one, or capture a screenshot.
# Called by guard.py; prints JSON. Never touches windows of unrelated programs.
# -Mode serve: stay running and answer one JSON request per stdin line ({"mode", "hwnd", "button", "out"})
# with one JSON line - starting PowerShell and compiling the types below costs ~0.5 s per call otherwise.
param(
    [ValidateSet('list', 'press', 'close', 'shot', 'serve')][string]$Mode = 'list',
    [long]$Hwnd = 0,
    [string]$Button = '',
    [string]$Out = ''
)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes, System.Drawing
Add-Type @"
using System; using System.Text; using System.Runtime.InteropServices; using System.Collections.Generic;
public class GuardWin {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc f, IntPtr l);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint p);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern bool IsWindowEnabled(IntPtr h);
  [DllImport("user32.dll")] public static extern bool IsHungAppWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern IntPtr GetWindow(IntPtr h, uint cmd);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
  [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint flags);
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [StructLayout(LayoutKind.Sequential)] public struct POINT { public int X, Y; }
  [DllImport("user32.dll")] public static extern bool ScreenToClient(IntPtr h, ref POINT p);
  public static List<IntPtr> Visible() { var r = new List<IntPtr>(); EnumWindows((h, l) => { if (IsWindowVisible(h)) r.Add(h); return true; }, IntPtr.Zero); return r; }
  public static string Text(IntPtr h) { var s = new StringBuilder(512); GetWindowText(h, s, 512); return s.ToString(); }
  public static string Cls(IntPtr h) { var s = new StringBuilder(256); GetClassName(h, s, 256); return s.ToString(); }
  public static uint Pid(IntPtr h) { uint p; GetWindowThreadProcessId(h, out p); return p; }
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr parent, EnumProc f, IntPtr l);
  // a Win32 child control (button) whose caption, without '&' accelerators, equals text
  public static IntPtr FindChild(IntPtr parent, string text) {
    IntPtr found = IntPtr.Zero;
    EnumChildWindows(parent, (h, l) => {
      if (Text(h).Replace("&", "").Trim().Equals(text, StringComparison.OrdinalIgnoreCase) && IsWindowVisible(h)) { found = h; return false; }
      return true; }, IntPtr.Zero);
    return found;
  }
}
"@
Add-Type -ReferencedAssemblies Accessibility @"
using System; using System.Runtime.InteropServices; using Accessibility;
public class GuardAcc {
  [StructLayout(LayoutKind.Sequential)] public struct PT { public int X, Y; }
  [DllImport("oleacc.dll")] static extern int AccessibleObjectFromPoint(PT pt, [MarshalAs(UnmanagedType.Interface)] out IAccessible acc, out object child);
  // MSAA default action (what screen readers do) of the element at a screen point
  public static string DoDefault(int x, int y) {
    IAccessible acc; object child; PT p; p.X = x; p.Y = y;
    if (AccessibleObjectFromPoint(p, out acc, out child) != 0 || acc == null) return null;
    string name = acc.get_accName(child);
    acc.accDoDefaultAction(child);
    return name;
  }
}
"@
$A = [System.Windows.Automation.AutomationElement]
$TS = [System.Windows.Automation.TreeScope]
# processes whose windows are always relevant even if not owned by AutoCAD
$ExtraProc = '^(OpenWith|senddmp|cer_dialog|cer.*ui|WerFault|AdskErrorReport.*)$'

function Get-Elements([IntPtr]$h, [int]$max = 60) {
    $out = @()
    try {
        $root = $A::FromHandle($h)
        foreach ($e in $root.FindAll($TS::Descendants, [System.Windows.Automation.Condition]::TrueCondition)) {
            $n = $e.Current.Name
            if (-not $n) { continue }
            $r = $e.Current.BoundingRectangle
            $out += [pscustomobject]@{ name = $n.Trim(); type = $e.Current.ControlType.ProgrammaticName -replace '^ControlType\.', ''; area = [int]($r.Width * $r.Height) }
            if ($out.Count -ge $max) { break }
        }
    } catch {}
    return $out
}

function Get-AcadWindows {
    $acad = @(Get-Process acad -ErrorAction SilentlyContinue)
    $pids = @($acad | ForEach-Object { [uint32]$_.Id })
    # AutoCAD's main frames by class: during start-up Process.MainWindowHandle can point elsewhere
    $visible = [GuardWin]::Visible()
    $mains = @($visible | Where-Object { ($pids -contains [GuardWin]::Pid($_)) -and ([GuardWin]::Cls($_) -like 'AfxMDIFrame*') })
    $names = @{}; Get-Process | ForEach-Object { $names[[uint32]$_.Id] = $_.ProcessName }
    $res = @()
    foreach ($h in $visible) {
        if ($mains -contains $h) { continue }
        $procId = [GuardWin]::Pid($h); $owner = [GuardWin]::GetWindow($h, 4)
        $pn = $names[$procId]
        if (-not (($pids -contains $procId) -or ($mains -contains $owner) -or ($pn -match $ExtraProc))) { continue }
        $res += [pscustomobject]@{
            hwnd = [long]$h; pid = $procId; process = $pn; class = [GuardWin]::Cls($h); title = [GuardWin]::Text($h)
            owned_by_autocad = ($mains -contains $owner); enabled = [GuardWin]::IsWindowEnabled($h); elements = @(Get-Elements $h)
        }
    }
    [pscustomobject]@{
        autocad_running = ($acad.Count -gt 0)
        autocad_pids = $pids
        autocad_blocked = [bool]($mains | Where-Object { -not [GuardWin]::IsWindowEnabled($_) })
        autocad_hung = [bool]($mains | Where-Object { [GuardWin]::IsHungAppWindow($_) })
        windows = @($res)
    }
}

function Press([IntPtr]$h, [string]$name) {
    $root = $A::FromHandle($h)
    $all = @($root.FindAll($TS::Descendants, [System.Windows.Automation.Condition]::TrueCondition))
    $cands = @($all | Where-Object { $_.Current.Name -and $_.Current.Name.Trim() -eq $name })
    if (-not $cands) { $cands = @($all | Where-Object { $_.Current.Name -and $_.Current.Name.Trim() -like "$name*" }) }
    if (-not $cands) { return 'not-found' }
    # real buttons first, then the smallest element (a group box titled 'Ignore' is bigger than the button)
    $el = $cands | Sort-Object @{ e = { if ($_.Current.ControlType.ProgrammaticName -eq 'ControlType.Button') { 0 } else { 1 } } },
                                @{ e = { $_.Current.BoundingRectangle.Width * $_.Current.BoundingRectangle.Height } } | Select-Object -First 1
    foreach ($pat in @([System.Windows.Automation.InvokePattern]::Pattern, [System.Windows.Automation.SelectionItemPattern]::Pattern, [System.Windows.Automation.TogglePattern]::Pattern)) {
        $o = $null
        if ($el.TryGetCurrentPattern($pat, [ref]$o)) {
            if ($o -is [System.Windows.Automation.InvokePattern]) { $o.Invoke(); return 'invoke' }
            if ($o -is [System.Windows.Automation.SelectionItemPattern]) { $o.Select(); return 'select' }
            if ($o -is [System.Windows.Automation.TogglePattern]) { $o.Toggle(); return 'toggle' }
        }
    }
    # Win32 dialog (#32770): the button is a child window with that caption
    $child = [GuardWin]::FindChild($h, $name)
    if ($child -ne [IntPtr]::Zero) {
        [GuardWin]::PostMessage($child, 0xF5, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null  # BM_CLICK
        return 'bm_click'
    }
    $r = $el.Current.BoundingRectangle
    # MSAA default action at the element's centre (task dialogs, DirectUI)
    try {
        $accName = [GuardAcc]::DoDefault([int]($r.X + $r.Width / 2), [int]($r.Y + $r.Height / 2))
        if ($accName -and $accName.Replace('&', '').Trim() -eq $name) { return 'msaa' }
    } catch {}
    # last resort (e.g. Qt widgets): post a mouse click to the window at the element's centre
    $p = New-Object GuardWin+POINT; $p.X = [int]($r.X + $r.Width / 2); $p.Y = [int]($r.Y + $r.Height / 2)
    [GuardWin]::ScreenToClient($h, [ref]$p) | Out-Null
    $lp = [IntPtr](($p.Y -shl 16) -bor ($p.X -band 0xFFFF))
    [GuardWin]::PostMessage($h, 0x201, [IntPtr]1, $lp) | Out-Null
    Start-Sleep -Milliseconds 80
    [GuardWin]::PostMessage($h, 0x202, [IntPtr]0, $lp) | Out-Null
    return 'click'
}

function Shot([IntPtr]$h, [string]$path) {
    $r = New-Object GuardWin+RECT
    [GuardWin]::GetWindowRect($h, [ref]$r) | Out-Null
    $w = [Math]::Max(1, $r.R - $r.L); $hh = [Math]::Max(1, $r.B - $r.T)
    $bmp = New-Object System.Drawing.Bitmap $w, $hh
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $dc = $g.GetHdc(); [GuardWin]::PrintWindow($h, $dc, 2) | Out-Null; $g.ReleaseHdc($dc)
    $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose(); $bmp.Dispose()
    return "$w x $hh"
}

# is the window still there after a press / close? (polled: most close within 100 ms)
function Still-Open([IntPtr]$h, [int]$ms = 700) {
    $t = [Diagnostics.Stopwatch]::StartNew()
    do {
        if (-not ([GuardWin]::IsWindow($h) -and [GuardWin]::IsWindowVisible($h))) { return $false }
        Start-Sleep -Milliseconds 40
    } while ($t.ElapsedMilliseconds -lt $ms)
    return ([GuardWin]::IsWindow($h) -and [GuardWin]::IsWindowVisible($h))
}

function Invoke-Mode([string]$mode, [long]$hwnd, [string]$button, [string]$out) {
    $h = [IntPtr]$hwnd
    switch ($mode) {
        'list' { return (Get-AcadWindows | ConvertTo-Json -Depth 5 -Compress) }
        'press' {
            $how = Press $h $button
            return ([pscustomobject]@{ hwnd = $hwnd; button = $button; method = $how; still_open = (Still-Open $h) } | ConvertTo-Json -Compress)
        }
        'close' {
            [GuardWin]::PostMessage($h, 0x10, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null
            return ([pscustomobject]@{ hwnd = $hwnd; method = 'WM_CLOSE'; still_open = (Still-Open $h) } | ConvertTo-Json -Compress)
        }
        'shot' {
            $size = Shot $h $out
            return ([pscustomobject]@{ hwnd = $hwnd; path = $out; size = $size } | ConvertTo-Json -Compress)
        }
    }
    throw "unknown mode '$mode'"
}

if ($Mode -eq 'serve') {
    [Console]::InputEncoding = [System.Text.Encoding]::UTF8
    while ($null -ne ($line = [Console]::In.ReadLine())) {   # ends when guard.py closes the pipe
        try {
            $q = $line | ConvertFrom-Json
            $r = Invoke-Mode ([string]$q.mode) ([long]$q.hwnd) ([string]$q.button) ([string]$q.out)
        } catch {
            $r = [pscustomobject]@{ error = "$_" } | ConvertTo-Json -Compress
        }
        [Console]::Out.WriteLine($r)
        [Console]::Out.Flush()
    }
} else {
    Invoke-Mode $Mode $Hwnd $Button $Out
}
