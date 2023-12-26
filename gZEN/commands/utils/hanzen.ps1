Param(
  [Parameter(Mandatory = $True, Position = 0)]
  [ValidatePattern("wide|narrow")]$Task,
  [Parameter(Mandatory = $True, Position = 1)]
  [String]$_str
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

function fn_To_Wide {
  Param(
    [Parameter(Mandatory = $True)][String]$half_string
  )
  Add-Type -AssemblyName "Microsoft.VisualBasic"
  [Microsoft.VisualBasic.Strings]::StrConv($half_string, [Microsoft.VisualBasic.VbStrConv]::Wide)
}

function fn_To_Narrow {
  Param(
    [Parameter(Mandatory = $True)][String]$wide_string
  )
  Add-Type -AssemblyName "Microsoft.VisualBasic"
  [Microsoft.VisualBasic.Strings]::StrConv($wide_string, [Microsoft.VisualBasic.VbStrConv]::Narrow)
}

switch ($Task) {
  wide { fn_To_Wide $_str}
  narrow { fn_To_Narrow $_str}
}
