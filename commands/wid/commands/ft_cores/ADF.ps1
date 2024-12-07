Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

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
  [PSCustomObject[]]$script:incomplete_list = [FT_Array]::Incomplete($main_list)

  [PSCustomObject[]]$script:diff_list = [FT_Array]::No_Duplicates($main_list, $_addition_list, $_target_key)

  # 中身が$null であるオブジェクトを関数のパラメータとして渡すことはできない。
  # Result型というものがもっと便利らしい

  # 差分なし、不完全なし
  if (($null -eq $diff_list) -and ($null -eq $incomplete_list)) {
    Write-Host './ft_cores/ADK: 追記できるものはありません。'
    return 0
  }
  
  # 差分あり、不完全なし
  if (($null -ne $diff_list) -and ($null -eq $incomplete_list)) {
    return $main_list + $diff_list
  }
  
  # 差分あり、不完全あり
  # $_main_list の既存のプロパティを $addition に上書き変更させてはならない
  if (($null -ne $diff_list) -and ($null -ne $incomplete_list)) {
    [PSCustomObject[]]$private:filled_arr = [FT_Array]::Filled($_addition_list, $incomplete_list, $_target_key)
    $result = if ($Null -ne $filled_arr) {
      # TODO: リファクタリンすること
      [PSCustomObject[]]$private:excluds = [FT_Array]::No_Duplicates($main_list, $filled_arr, $_target_key)
      return ($excluds + $filled_arr + $diff_list)
    }
    else { 
      return ($main_list + $diff_list)
    }
    return $result
  }
  

  # 差分なし、不完全あり
  if (($null -eq $diff_list) -and ($null -ne $incomplete_list)) {
    [PSCustomObject[]]$private:filled_arr = [FT_Array]::Filled($_addition_list, $incomplete_list, $_target_key)
    $result = if ($Null -ne $filled_arr) {
      [PSCustomObject[]]$private:excluds = [FT_Array]::No_Duplicates($main_list, $filled_arr, $_target_key)
      return ($excluds + $filled_arr)
    }
    else {
      Write-Host '追記できるものはありません。'
      return 0
    }
    return $result
  }
}

