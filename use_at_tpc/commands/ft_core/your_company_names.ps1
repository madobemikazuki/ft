Set-StrictMode -Version 3.0

Set-Variable -Name Z_SLASH -Value " / " -Option Constant

function your_company_names {
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory)][String]$managemanet_com_name,
    [Parameter(Mandatory)][String]$employer_name
  )
  
  if ($managemanet_com_name -eq $employer_name) {
    return $managemanet_com_name
  }

  # 二つの名前が違うとき実行
  if (!($managemanet_com_name -eq $employer_name)){
    . .\ft_core\combined_name.ps1
    return combined_name $managemanet_com_name $employer_name $Z_SLASH
  }
}
exit 0
