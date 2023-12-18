Param(
  [Parameter(Mandatory = $true)]
  [String]$dic,
  [Parameter(Mandatory = $true)]
  [String]$target
)

<#
.SYNOPSIS
  指定したフォルダパス以下のファイルから $target に合致するものを抽出する
.DESCRIPTION
  
.EXAMPLE

.EXAMPLE
.INPUTS
.NOTES

#>

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

# ファイルが存在しないときのエラー処理がわからない。
try {
  $file_name = Get-childItem -Path $dic -Name -File -Include $target
  return $file_name
}
catch {
  Write-Host "エラー発生 :: $($_.Exception.Message)"
}