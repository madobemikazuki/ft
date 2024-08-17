Param(
  [Parameter(Mandatory = $True, Position = 0)]
  [ValidatePattern("^\d{8}")]
  [String]$_date,
  [Parameter(Mandatory = $True, Position = 1)]
  [ValidatePattern("^\d{6}")]
  [String]$_wid,
  [Parameter(Mandatory = $True, Position = 2)]
  [ValidatePattern("^\d{2}\b\-\b\d{6}$")]
  [ValidateCount(1, 10)]
  [String[]]$_central_nums
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

function fn_Read {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [ValidatePattern("\.csv$|\.json$")]$_path
  )
  switch -Regex ($_path) {
    "\.csv$" {
      return Import-Csv -Path $_path -Encoding Default
    }
    "\.json$" {
      return Get-Content -Path $_path -Encoding UTF8 | ConvertFrom-Json
    }
    Default {
      Write-Host "拡張子が該当しないので終了。"
      exit 0
    }
  }
}

# おそらく不要な関数
function fn_Search_WID {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [PSCustomObject[]]$_wid_list,
    [Parameter(Mandatory = $True, Position = 1)]
    [ValidatePattern("^\d{6}")]
    [String]$_target_wid_num
  )
  $LookUpHash = @{}
  foreach ($_wid in $_wid_list) {
    $LookUpHash[$_wid."作業件名コード"] = $_wid
  }
  return $LookUpHash.$_target_wid_num."作業主管グループ"
}

function fn_Slash_Format {
  Param(
    [Parameter(Mandatory)]
    [ValidatePattern("^\d{8}")]
    [String]$_date
  )
  . .\ft_core\Excel-Hell-Format.ps1
  Slash $_date
}

function fn_Search_Registerer {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [PSCustomObject]$_Obj,
    [Parameter(Mandatory = $True, Position = 1)]
    [ValidatePattern("^\d{2}\b\-\b\d{6}$")]
    [String[]]$_register_num_list
  )
  [PSCustomObject[]]$list = foreach ($_ in $_register_num_list) {
    $_Obj.$_
  }
  return $list
}

function script:fn_Shorten_Com_Type_Name {
  Param(
    [Parameter(Mandatory = $True)]
    [String]$_corporate_name
  )
  switch ($_corporate_name) {
    { $_ -match "株式会社|\(株\)" } { return $_ -replace "株式会社|\(株\)", "（株）" }
    { $_ -match "有限会社|\(有\)" } { return $_ -replace "有限会社|\(有\)", "（有）" }
    default { return $_ }
  }
}

function script:fn_To_Wide {
  Param(
    [Parameter(Mandatory = $True)][String]$half_string
  )
  Add-Type -AssemblyName "Microsoft.VisualBasic"
  [Microsoft.VisualBasic.Strings]::StrConv($half_string, [Microsoft.VisualBasic.VbStrConv]::Wide)
}

function script:fn_WBC_Company_Names {
  Param(
    [Parameter(mandatory = $True, Position = 0)]
    [PSCustomObject][ref]$_applicant
  )
  # この処理は bind_r.ps1 とは若干違う。
  # 理由: 登録者管理リスト_coh.csv のフィールドと事前申請のフィールドが異なるため。
  # if で判定する前に企業名を fn_Shorten_Com_Type_Nameするのは
  # ソースとなる Companies.json に半角全角の(株)（株）が混在しているためである。
  $application_com = fn_Shorten_Com_Type_Name $_applicant.'電力申請会社名称' 
  $manegement_com = fn_Shorten_Com_Type_Name $_applicant.'管理会社名称'
  $employment_com = fn_Shorten_Com_Type_Name $_applicant.'雇用名称'

  if ($manegement_com.Contains("派遣")) { return $manegement_com }
  if ($application_com -eq $employment_com) { return $application_com }
  return ($application_com + "／" + $employment_com)
}

function script:fn_Combined_Name {
  Param(
    [Parameter(Mandatory = $true, Position = 0)][String]$first_name,
    [Parameter(Mandatory = $true, Position = 1)][String]$last_name,
    [String]$delimiter = '　'#デフォルト引数 呼び出し側で -delimiter を指定すること
  )
  $sb = New-Object System.Text.StringBuilder
  #副作用処理  StringBuilderならちょっと速いらしい。要素数が少ないから意味ないかも。
  @($first_name, $delimiter , $last_name) | ForEach-Object { [void] $sb.Append($_) }
  return $sb.ToString()
}

function script:fn_DisQualification {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [PSCustomObject[]][ref]$_obj_list,
    [Parameter(Mandatory = $True, Position = 1)]
    [String[]][ref]$_prequisites
  )
  # 必須条件としての値が一意であること。
  $boolean_list = foreach ($_preq in $_prequisites) {
    # @を省くと自動的に単一のObject が返ってくる。
    @($_obj_list | Select-Object -Property $_preq | Sort-Object -Property $_preq -Unique).length
  }
  $result = foreach ($_ in $boolean_list) { $_ -eq 1 }
  Write-Host $result
  return ($result -contains $False)
}


function fn_Transform_Cancelation {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [PSCustomObject[]]$_obj_list
  )
  $list = foreach ($_ in $_obj_list) {
    $obj = [ordered]@{}
    $obj["漢字氏名"] = fn_Combined_Name $_."氏名（姓）" $_."氏名（名）"
    $obj["氏名（カナ）"] = fn_To_Wide $_."氏名（カナ）"
    $obj["電力申請会社番号"] = ($_."電力申請会社番号").replace("T", "0")
    $obj["電力申請会社名称"] = fn_Shorten_Com_Type_Name $_."電力申請会社名称"
    $obj["作業者証番号"] = [String]($_."作業者証番号")
    $obj["東電管理番号"] = [String]($_."東電管理番号")
    $obj["解除WBC_申請会社"] = fn_WBC_Company_Names ([ref]$_)
    
    # TODO: $config.json の c_selection に漢字氏名、解除WBC_申請会社を追加する。
    # TODO: 漢字氏名 追加する
    # TODO: 解除WBC_申請会社 追加する 会社形態の省略文字の適用
    # TODO: 電力申請会社番号 の頭文字 T を 0 に変更する
    # TODO: 電力申請会社名称の 会社形態の省略文字の適用
    [PSCustomObject]$obj
  }
  return $list
}

#if($_central_nums.length -lt 11){}

$config = fn_Read ".\config\cnc.json"

$script:reserved_date = fn_Slash_Format $_date
Write-Host $reserved_date

$wid_path = (${HOME} + $config.wid_path)
$wid_LookUpHash = fn_Read $wid_path
try { $script:wid = $wid_LookUpHash.$_wid }
catch {
  Start-Process notepad.exe $wid_path
  Throw "Error :指定した WID番号 には主管グループ名が設定されていません💩"
}

remove-variable wid_LookUpHash

Write-Host $wid.depertment
Write-Host $wid.group

$script:poe_config = $config.poe_config

$common_obj = [PSCustomObject]@{
  "解除予約日"   = $reserved_date
  "担当主管課班名" = @($wid.depertment, $wid.group) -join "`r`n"
}
$common_address = $poe_config.common_address_table
#$common_obj | Format-Table
$common_printing_obj = . .\ft_core\POE\core\common_posting_format.ps1 $common_obj $common_address
$common_printing_obj | Format-Table


# 該当者を検索
$registerer_obj = fn_Read (${HOME} + $config.registerer_path)
#Write-Host "ここおかしくね？"
#Write-host $registerer_obj.gettype()
$registerer_list = fn_Search_Registerer $registerer_obj $_central_nums
$registerer_list | Format-Table
if (fn_DisQualification ([ref]$registerer_list) ([ref]$config.prequisites)) {
  throw "申請者全員の申請会社名称、もしくは雇用名称が異なりますね。💩"
  exit 0
}

remove-variable registerer_obj


# フィールドを絞りこむ
$extraction_list = $config.extraction_list
$extracted_list = $registerer_list | Select-Object -Property $extraction_list
remove-variable registerer_list


# 必要な情報に整形する
$transformed_list = fn_Transform_Cancelation $extracted_list
remove-variable extracted_list
$address_table = $poe_config.address_table
$main_obj_list = . .\ft_core\POE\core\list_posting_format.ps1 $transformed_list $address_table


# 転記情報を統合する
[PSCustomObject[]]$integrated_list = $common_printing_obj + $main_obj_list

$integrated_list | Format-Table
#Write-Host $integrated_list.GetType()
#$poe_config| Format-List
# POE へ
. .\ft_core\POE\poe.ps1 $poe_config $integrated_list

