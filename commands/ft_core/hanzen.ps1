Set-StrictMode -Version 3.0

function to_wide {
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory)][String]$half_string
  )
  Add-Type -AssemblyName "Microsoft.VisualBasic"
  [Microsoft.VisualBasic.Strings]::StrConv($half_string, [Microsoft.VisualBasic.VbStrConv]::Wide)
}

function to_narrow {
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory)][String]$wide_string
  )
  Add-Type -AssemblyName "Microsoft.VisualBasic"
  [Microsoft.VisualBasic.Strings]::StrConv($wide_string, [Microsoft.VisualBasic.VbStrConv]::Narrow)
}

exit 0