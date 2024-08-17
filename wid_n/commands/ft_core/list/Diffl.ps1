function Diffl {
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [PSCustomObject[]]$_exists_list,
    [Parameter(Mandatory = $True, Position = 1)]
    [PSCustomObject[]]$_addition_list,
    [Parameter(Mandatory = $True, Position = 2)]
    [String]$_target_key
  )
  Set-StrictMode -Version 3.0
  $ErrorActionPreference = "Stop"
  # 二つの配列を受け取り、追加側リストに既存リストの特定のプロパティの値と重複しないオブジェクトの配列を返す。

  $private:result = $_addition_list | Where-Object{$_.$_target_key -notin ($_exists_list).$_target_key}
  return $result
  
  <#
  $private:exists_values = foreach ($_ in $_exists_list) { 
    $_.$_target_key
  }

  $private:diff_list = $_addition_list | Where-Object { $exists_values -notcontains $_.$_target_key }
  return $diff_list
  #>
}

