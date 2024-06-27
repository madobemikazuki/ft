Param(
  [Parameter(Mandatory = $True, Position = 0)]
  [PSCustomObject[]]$_arr,
  [Parameter(Mandatory = $True, Position = 1)]
  [int16]$_range
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

#配列を任意の$_range を一単位として分割する。
if ($_range -lt $_arr.Length) {
  $private:new_arr = @()
    
  $private:part = [Math]::Floor($_arr.length / $_range)
  Write-Host $part " : 除算した数(少数点以下を切捨て)" 
    
  $private:remainder = $_arr.length % $_range
  Write-Host $remainder " : 除算して余った数"

  $start_index = 0
  foreach ($_ in @(1..$part)) {
    #Write-Host $arr($start_index..-1).length
    $split_Item = $_arr[$start_index..($start_index + $_range - 1)]
    $start_index += $_range
    $new_arr += , [PSCustomObject[]]$split_Item
  }
  if ($remainder -gt 0) {
    # $remainder がゼロより大きければ端数の配列をジャグ配列として返す
    $new_arr += , [PSCustomObject[]]$_arr[-1.. - $remainder]
  }
  return $new_arr
}
exit 0
