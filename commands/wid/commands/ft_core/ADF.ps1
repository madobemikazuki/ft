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


function script:fn_Integrate {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [PSCustomObject[]][ref]$_main_list,
    [Parameter(Mandatory = $True, Position = 1)]
    [PSCustomObject[]][ref]$_filled_list,
    [Parameter(Mandatory = $True, Position = 2)]
    [PSCustomObject[]][ref]$_diff_list,
    [Parameter(Mandatory = $True, Position = 3)]
    [String][ref]$_target_key
  )
  . .\ft_core\list\Excluded-List.ps1
  [PSCustomObject[]]$private:excludeds = Excluded-List $_main_list $_filled_list $_target_key
  [PSCustomObject[]]$private:integrated_list = ($excludeds + $_filled_list + $_diff_list)
  #Write-Host "最終リストの要素数 : " $integrated_list.Length
  return $integrated_list
}


function ADF {
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [PSCustomObject[]]$_main_list,
    [Parameter(Mandatory = $True, Position = 1)]
    [PSCustomObject[]]$_addition_list,
    [Parameter(Mandatory = $True, Position = 2)]
    [String]$_target_key
  )

  $script:main_list = $_main_list
  . .\ft_core\list\Incomplete-List.ps1
  [PSCustomObject[]]$script:incomplete_list = Incomplete-List $main_list

  . .\ft_core\list\Diffl.ps1
  [PSCustomObject[]]$script:diff_list = Diffl $main_list $_addition_list $_target_key

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
    . .\ft_core\list\Filled-List.ps1
    [PSCustomObject[]]$private:filled_list = Filled-List $_addition_list $incomplete_list $_target_key

    $result = if ($Null -ne $filled_list) {
      Write-Host ""
      . .\ft_core\list\Excluded-List.ps1
      [PSCustomObject[]]$private:excluds = Excluded-List $main_list $filled_list $_target_key
      return ($excluds + $filled_list + $diff_list)
    }
    else { return ($main_list + $diff_list) }
    
    return $result
  }
  

  # 差分なし、不完全あり
  if (($null -eq $diff_list) -and ($null -ne $incomplete_list)) {
    . .\ft_core\list\Filled-List.ps1
    [PSCustomObject[]]$private:filled_list = Filled-List $_addition_list $incomplete_list $_target_key
    . .\ft_core\list\Excluded-List.ps1
    [PSCustomObject[]]$private:excluds = Excluded-List $main_list $filled_list $_target_key
    [PSCustomObject[]]$private:complete_list = $excluds + $filled_list
    return $complete_list
  }
}

