Param(
  [Parameter(Mandatory = $True, Position = 0)][String]$_managemanet_com_name,
  [Parameter(Mandatory = $True, Position = 1)][String]$_employer_name
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

if ($_managemanet_com_name -eq $_employer_name) {
  return $_managemanet_com_name
}

# 二つの名前が違うとき実行
if (!($_managemanet_com_name -eq $_employer_name)) {
  return . .\ft_core\combined_name.ps1 $_managemanet_com_name $_employer_name  -delimiter " / "
}
