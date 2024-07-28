Param(
  [Parameter(Mandatory = $True, Position = 0)]
  [PSCustomObject[]]$_main_list,
  [Parameter(Mandatory = $True, Position = 1)]
  [PSCustomObject[]]$_addition_list,
  [Parameter(Mandatory = $True, Position = 2)]
  [String]$_target_key
)
Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

function script:fn_Exists {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [PSCustomObject[]]$_obj_list
  )
  return ($null -eq $_obj_list)

}

function script:fn_Not_Exists {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [PSCustomObject[]]$_obj_list
  )
  return ($null -ne $_obj_list)
}



  $script:main_list = $_main_list
  [PSCustomObject[]]$script:incomplete_list = . .\ft_core\list\Incomplete-List.ps1 $main_list
  [PSCustomObject[]]$script:diff_list = . .\ft_core\list\Diffl.ps1 $main_list $_addition_list $_target_key

  # 中身が$null であるオブジェクトを関数のパラメータとして渡すことはできない。
  # Result型というものがもっと便利らしい

  # 差分なし、不完全なし
  if (($null -eq $diff_list) -and ($null -eq $incomplete_list)) {
    Write-Host '追記できるものはありません。'
    return 0
  }
    
  if (($null -ne $diff_list) -and ($null -eq $incomplete_list)) {
    return $main_list + $diff_list
  }
  
  # 差分あり、不完全あり
  # $_main_list の既存のプロパティを $addition に上書き変更させてはならない
  if (($null -ne $diff_list) -and ($null -ne $incomplete_list)) {
    [PSCustomObject[]]$private:filled_list = . .\ft_core\list\Filled-List.ps1 $_addition_list $incomplete_list $_target_key
    $result = if ($Null -ne $filled_list) {
      Write-Host ""
      [PSCustomObject[]]$private:excluds = . .\ft_core\list\Excluded-List.ps1 $main_list $filled_list $_target_key
      return ($excluds + $filled_list + $diff_list)
    }
    else { return ($main_list + $diff_list) }
    
    return $result
  }
  

  # 差分なし、不完全あり
  if (($null -eq $diff_list) -and ($null -ne $incomplete_list)) {
    
    [PSCustomObject[]]$private:filled_list = . .\ft_core\list\Filled-List.ps1 $_addition_list $incomplete_list $_target_key
    [PSCustomObject[]]$private:excluds = . .\ft_core\list\Excluded-List.ps1 $main_list $filled_list $_target_key
    [PSCustomObject[]]$private:complete_list = $excluds + $filled_list
    return $complete_list
  }
