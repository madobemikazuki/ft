Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"


<#
  各種実務コマンド用のデータ抽出・整形用の具体的な実装の関数群
#>

. ..\ft_cores\FT_IO.ps1;
. ..\ft_cores\FT_Name.ps1
. ..\ft_cores\Odd_Name.ps1
. ..\ft_cores\FT_Array.ps1
. ..\ft_cores\FT_Object.ps1


function fn_SF_Read_Config {
  Param(  
    [Parameter(Mandatory = $True, Position = 0)]
    [String]$_path, 
    [Parameter(Mandatory = $True, Position = 1)]
    [String]$_field
  )
  return ([FT_IO]::Read_JSON_Object($_path)).$_field
}

function fn_SF_Read_as_Dict {
  Param(  
    [Parameter(Mandatory = $True, Position = 0)]
    [String]$_path, 
    [Parameter(Mandatory = $True, Position = 1)]
    [String]$_primary_key
  )
  $private:candidates_arr = [FT_IO]::Read_JSON_Array((${HOME} + $_path))
  $dict = [FT_Array]::ToDict($candidates_arr, $_primary_key)
  return $dict 
}

function fn_SF_Write_JSON_Array {
  Param(  
    [Parameter(Mandatory = $True, Position = 0)]
    [String]$_path,
    [Parameter(Mandatory = $True, Position = 1)]
    [PSCustomObject[]]$_arr
  )
  [FT_IO]::Write_JSON_Array((${Home} + $_path), $_arr)
}

function script:fn_SF_R_Company_Names {
  Param(
    [Parameter(mandatory = $True, Position = 0)]
    [PSCustomObject]$_applicant
  )
  $application_com = [FT_Name]::Shortened_Com_Type_Name($_applicant.'FT_所属企業名')
  $manegement_com = if ([String]::IsNullOrEmpty($_applicant.'管理会社')) {
    [FT_Name]::Shortened_Com_Type_Name($_applicant.'FT_所属企業名')
  }
  else { [FT_Name]::Shortened_Com_Type_Name($_applicant.'管理会社') }

  $employment_com = if ([String]::IsNullOrEmpty($_applicant.'雇用企業名称（漢字）')) {
    [FT_Name]::Shortened_Com_Type_Name($_applicant.'FT_所属企業名') 
  }
  else { [FT_Name]::Shortened_Com_Type_Name($_applicant.'雇用企業名称（漢字）') }
  
  if ($manegement_com.Contains("派遣")) { return $manegement_com + '／' + $employment_com }
  if ($application_com -eq $employment_com) { return $application_com }
  if ($application_com -ne $employment_com) { return ($application_com + '／' + $employment_com) }
}

function script:fn_SF_To_Application_Number {
  Param(
    [Parameter(mandatory = $True, Position = 0)]
    [String]$_n
  )
  if ($_n -match "^T") {
    return $_n.Replace('T', '0')
  }
  else {
    Throw ($_n + " は 頭文字 T ではないため処理できません。")
  }
}

function script:fn_SF_C_Company_Names {
  Param(
    [Parameter(mandatory = $True, Position = 0)]
    [PSCustomObject]$_applicant
  )
  $application_com = [FT_Name]::Shortened_Com_Type_Name($_applicant.'電力申請会社名称')
  $manegement_com = [FT_Name]::Shortened_Com_Type_Name($_applicant.'管理会社名称')
  $employment_com = if ([String]::IsNullOrEmpty($_applicant.'雇用名称')) {
    [FT_Name]::Shortened_Com_Type_Name($_applicant.'電力申請会社名称') 
  }
  else { [FT_Name]::Shortened_Com_Type_Name($_applicant.'雇用名称') }
  
  if ($manegement_com.Contains("派遣")) { return $manegement_com }
  if ($application_com -eq $employment_com) { return $application_com }
  if ($application_com -ne $employment_com) { return ($application_com + '／' + $employment_com) }
}

function fn_SF_Marge_Dicts {  
  # $_candidates_dict に $_reserved_dict を合成してPSCustomObject[]として返す
  param (
    [Parameter(mandatory = $True, Position = 0)]
    [HashTable]$_candidates_dict,
    [Parameter(mandatory = $True, Position = 1)]
    [HashTable]$_reserved_dict
  )
  $marged_arr = foreach ($_id in $_candidates_dict.keys) {
    $candidates = $_candidates_dict.$_id
    $reserved = if ($_reserved_dict.ContainsKey($_id)) {
      $_reserved_dict.$_id
    }
    else { continue }
    [FT_Object]::Marge($candidates, $reserved)
  }
  return $marged_arr
}

function fn_SF_Registration_Format {
  Param(
    [Parameter(mandatory = $True, Position = 0)]
    [PSCustomObject[]]$_marged_arr,
    [Parameter(Mandatory = $True, Position = 1)]
    [PSCustomObject]$odd    
  )
  Write-Host "登録モード"
  $completed_arr = foreach ($_ in $_marged_arr ) {
    # 登録モード用の関数を使用
    $wbc_app_com_names = fn_SF_R_Company_Names $_
    # Oddパターンをここで処理する。JSON側でパターンを定義する。
    $ed_com_name = [Odd_Name]::To_Appropriate((fn_SF_R_Company_Names $_), $odd)
    Add-Member -InputObject $_ -NotePropertyMembers @{
      $t_config.wbc_key = $wbc_app_com_names;
      $t_config.ed_key  = $ed_com_name;
    } -Force
    $_
  }
  $extracted_arr = [FT_Array]::Map($completed_arr, $t_config.extraction_target)
  $sorted_arr = [FT_Array]::Sort($extracted_arr, $t_config.addition_keys[0])
  Remove-Variable completed_arr, extracted_arr
  return $sorted_arr
}


function fn_SF_Cancellation_Format {
  Param(
    [Parameter(mandatory = $True, Position = 0)]
    [PSCustomObject[]]$_marged_arr,
    [Parameter(Mandatory = $True, Position = 1)]
    [PSCustomObject]$odd    
  )

  Write-Host "解除モード"
  $completed_arr = foreach ($_ in $_marged_arr ) {
    #TODO:Oddパターンをここで処理する。JSON側でパターンを定義する。
    Add-Member -InputObject $_ -NotePropertyMembers @{
      '漢字氏名（姓）'         = $_.'氏名（姓）';
      '漢字氏名（名）'         = $_.'氏名（名）';
      'FT_登録時申請会社番号'    = fn_SF_To_Application_Number $_.'電力申請会社番号';
      'FT_登録時申請会社名称'    = [FT_Name]::Shortened_Com_Type_Name($_.'電力申請会社名称');
      $t_config.wbc_key = fn_SF_C_Company_Names $_;
    } -Force
    $_
  }
  $extracted_arr = [FT_Array]::Map($completed_arr, $t_config.extraction_target)
  $sorted_arr = [FT_Array]::Sort($extracted_arr, $t_config.addition_keys[1])
  Remove-Variable completed_arr, extracted_arr
  return $sorted_arr
}  

