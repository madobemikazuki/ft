Param(
  [Parameter(Mandatory = $True, Position = 0)]
  [String[]]$header,
  [Parameter(Mandatory = $True, Position = 1)]
  [Object[]]$values,
  [String]$delimiter = ','
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

<#
.SYNOPSIS

.DESCRIPTION

.EXAMPLE
  
.EXAMPLE

.INPUTS

.NOTES

#>

<#
[PSCustomObject[]]$csv_object = $values | ConvertFrom-Csv -Header $header
return $csv_object
#>
$values | ConvertFrom-Csv -Header $header -Delimiter $delimiter
