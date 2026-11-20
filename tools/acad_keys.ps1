# Type into AutoCAD's command line with real keystrokes: bring AutoCAD to the front, verify that one of its
# windows IS the foreground window (the main frame or its floating command line, which keeps the focus
# while a command waits for input), then SendKeys. Refuses to type if another program keeps the focus.
# SendKeys syntax: {ENTER}, {ESC}; "+" "^" "%" "~" "(" ")" "{" "}" must be escaped in braces.
param([string]$Keys)
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8   # guard reads UTF-8
Add-Type -AssemblyName System.Windows.Forms
Add-Type @"
using System; using System.Runtime.InteropServices;
public class FG {
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int c);
  [DllImport("user32.dll")] public static extern void keybd_event(byte v, byte s, uint f, UIntPtr e);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint p);
  public static uint ForegroundPid() { uint p; GetWindowThreadProcessId(GetForegroundWindow(), out p); return p; }
}
"@
$acad = Get-Process acad | Select-Object -First 1
$main = $acad.MainWindowHandle
for ($i = 0; $i -lt 5 -and [FG]::ForegroundPid() -ne $acad.Id; $i++) {
  [FG]::keybd_event(0x12, 0, 0, [UIntPtr]::Zero); [FG]::keybd_event(0x12, 0, 2, [UIntPtr]::Zero)  # Alt tap unlocks SetForegroundWindow
  [FG]::SetForegroundWindow($main) | Out-Null
  Start-Sleep -Milliseconds 300
}
if ([FG]::ForegroundPid() -ne $acad.Id) { "AutoCAD is not in the foreground - nothing typed"; exit 1 }
[System.Windows.Forms.SendKeys]::SendWait($Keys)
"typed: $Keys"
