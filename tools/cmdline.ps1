# Read-only: text of AutoCAD's command line (prompt + recent history) via UI Automation.
param([int]$Tail = 15)
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8   # guard.command_line reads UTF-8 (cyrillic prompts)
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$A = [System.Windows.Automation.AutomationElement]
$p = Get-Process acad | Select-Object -First 1
$cond = New-Object System.Windows.Automation.PropertyCondition($A::ProcessIdProperty, $p.Id)
$lines = New-Object System.Collections.Generic.List[string]
foreach ($t in $A::RootElement.FindAll([System.Windows.Automation.TreeScope]::Children, $cond)) {
  $cli = $t.FindFirst([System.Windows.Automation.TreeScope]::Descendants,
    (New-Object System.Windows.Automation.PropertyCondition($A::ClassNameProperty, 'CommandLineControl')))
  if (-not $cli) { continue }
  foreach ($e in $cli.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition)) {
    $txt = $e.Current.Name
    if (-not $txt) {
      try { $txt = $e.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern).Current.Value } catch {}
    }
    if (-not $txt) {
      try { $txt = $e.GetCurrentPattern([System.Windows.Automation.TextPattern]::Pattern).DocumentRange.GetText(-1) } catch {}
    }
    if ($txt) { foreach ($l in ($txt -split "`r?`n")) { if ($l.Trim()) { $lines.Add($l.TrimEnd()) } } }
  }
}
if ($lines.Count -eq 0) { "no command-line text found" } else { $lines | Select-Object -Last $Tail }
