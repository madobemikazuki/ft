Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

function ToWide {
  Param(
    [Parameter(Mandatory = $True)][String]$half_string
  )
  Add-Type -AssemblyName "Microsoft.VisualBasic"
  [Microsoft.VisualBasic.Strings]::StrConv($half_string, [Microsoft.VisualBasic.VbStrConv]::Wide)
}

