Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

function ToWide {
  Param(
    [Parameter(Mandatory = $True)]
    [String]$_string
  )
  Add-Type -AssemblyName "Microsoft.VisualBasic"
  [Microsoft.VisualBasic.Strings]::StrConv($_string, [Microsoft.VisualBasic.VbStrConv]::Wide)
}

