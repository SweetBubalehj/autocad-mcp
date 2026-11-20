# Capture the AutoCAD main window even if other windows cover it (PrintWindow), without changing focus.
param([string]$Out)
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System; using System.Runtime.InteropServices;
public class PW {
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint flags);
}
"@
$p = Get-Process acad | Select-Object -First 1
$h = $p.MainWindowHandle
$r = New-Object PW+RECT
[PW]::GetWindowRect($h, [ref]$r) | Out-Null
$w = $r.R - $r.L; $hgt = $r.B - $r.T
$bmp = New-Object System.Drawing.Bitmap $w, $hgt
$g = [System.Drawing.Graphics]::FromImage($bmp)
$dc = $g.GetHdc(); [PW]::PrintWindow($h, $dc, 2) | Out-Null; $g.ReleaseHdc($dc)
$scale = [Math]::Min(1.0, 1600 / $w)
$small = New-Object System.Drawing.Bitmap ([int]($w * $scale)), ([int]($hgt * $scale))
$g2 = [System.Drawing.Graphics]::FromImage($small); $g2.InterpolationMode = 'HighQualityBicubic'
$g2.DrawImage($bmp, 0, 0, $small.Width, $small.Height)
$small.Save($Out, [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose(); $g2.Dispose(); $bmp.Dispose(); $small.Dispose()
"saved $Out ($w x $hgt)"
