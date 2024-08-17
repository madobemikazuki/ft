function script:fn_Contains_Empty {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [PSCustomObject]$_obj
  )
  $boolean_list = foreach ($_value in $_obj.psobject.Properties.Value) {
    [String]::IsNullOrEmpty($_value)
  }
  return ($True -in $boolean_list)
}

function Filled-List {
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [PSCustomObject[]]$_addition_list,
    [Parameter(Mandatory = $True, Position = 1)]
    [PSCustomObject[]]$_incomplete_list,
    [Parameter(Mandatory = $True, Position = 2)]
    [String]$_target_key
  )
  Set-StrictMode -Version 3.0
  $ErrorActionPreference = "Stop"
  $private:incomplete_values = foreach ($_ in $_incomplete_list) { $_.$_target_key }
  # 文字列リストの要素にリスト要素の特定のプロパティの値と一致するもので
  # かつ、当該プロパティの値に $null や '' が含まれていない要素を配列にして返す

  [PSCustomObject[]]$private:filled_list = $_addition_list | Where-Object {
    #追加リストの要素に不完全リストのプロパティが含まれていること
    ($incomplete_values.Contains($_.$_target_key) -and !(fn_Contains_Empty $_))
  }
    

  return $filled_list
}

