Param(
  [Parameter(Mandatory = $true, Position = 0)]
  [ValidatePattern('\.txt$')]$txt_path,
  [String]$encode = "Default"
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"
<#
.SYNOPSIS
  指定したフォルダパス以下のファイルから $target に合致するものを抽出する
.DESCRIPTION
  
.EXAMPLE

.EXAMPLE
.INPUTS
.NOTES

#>
return Get-Content -Path $txt_path  -Encoding $encode 
