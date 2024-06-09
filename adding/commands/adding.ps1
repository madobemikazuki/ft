Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

function script:fn_Export_JSON {
  Param(
    [Parameter(Mandatory = $true, Position = 0)][Object[]]$_obj,
    [Parameter(Mandatory = $true, Position = 1)][String]$_path
  )
  $utf8_with_BOM = New-Object System.Text.UTF8Encoding $True
  [System.IO.File]::WriteAllLines($_path, (ConvertTo-Json $_obj), $utf8_with_BOM)
}

function script:fn_Export_CSV {
  Param(
    [Parameter(Mandatory = $true, Position = 0)][Object[]]$_obj_list,
    [Parameter(Mandatory = $true, Position = 1)][String]$_path,
    [String]$_delimiter = ',',
    [String]$_encode = "utf8"# Default ではブラウザで参照すると文字化けする。
  )
  return $_obj_list | Export-Csv -NotypeInformation -Path $_path -Delimiter $_delimiter -Encoding $_encode -Force
}

function private:fn_Select_Sort_Unique{
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [Object[]]$_obj_list,
    [Parameter(Mandatory = $True, Position = 1)]
    [String[]]$_integrated_field,
    [Parameter(Mandatory = $True, Position = 2)]
    [String]$_extractor
  )
  $first = $_obj_list | Select-Object -Property $_integrated_field
  $second = $first | Sort-Object -Property $extractor -Unique
  return $second
}

function private:fn_Sub_Array {
  param (
    [Parameter(Mandatory =$True, Position=0)]
    [String[]]$_arr
  )
  if($null -eq $_arr){exit 404}
  return $_arr[1..($_arr.Length - 1)]
}

function private:fn_Head{
  Param (
    [Parameter(Mandatory = $true, Position = 0)]
    [String[]]$_arr
  )
  if($null -eq $_arr){exit 404}
  return $_arr[0]
}

function private:fn_Is_Multiple{
  Param (
      [Parameter(Mandatory = $true, Position = 0)]
      [String[]]$_list,
      [Parameter(Mandatory = $True, Position = 1)]
      [String]$_things
  )
  if($_list.Length -eq 1){
    Write-Host ($_things + "が複数存在しないので終了します。")
    exit 0
 }
}

function private:fn_Find_Files {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [String]$_Folder,
    [Parameter(Mandatory = $True, Position = 1)]
    [String[]]$_targets
  )
  $private:filename_list = Get-ChildItem -Path $_Folder -File -Name -Include $_targets
  if ($null -eq $filename_list) {
    fn_notifycation "エラー : $_targets ファイルが存在しません"
    exit 404
  }
  return $filename_list
}

function private:fn_Read {
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



$script:config = fn_Read ".\config\adding.json"
$script:source_folder = (${HOME} + $config.source_folder)

[String[]]$file_name_list = fn_Find_Files $source_folder $config.target_files
[String[]]$file_path_list = foreach ($_ in $file_name_list) { $source_folder + $_ }
fn_Is_Multiple $file_path_list "CSVファイル"

# 主体となるオブジェクトの指定
# 基本的に取得したファイルパスリストの先頭要素とする。
# 面倒なconfig指定を避けるため。
$main_path = fn_head $file_path_list
$main_csv = fn_Read $main_path
$sub_csv_list = fn_Sub_Array $file_path_list

$extractor = $config.extractor
$new_csv = [PSCustomObject[]]@()

# $main_csv に追記すべく sub_csv_list から各要素のプロパティを転写していく
foreach ($_file in $sub_csv_list) {
  $csv = fn_Read $_file
  #$csv|Format-Table
  foreach($_obj in $csv){
    $main_obj = $main_csv | Where-Object {$_.$extractor -eq $_obj.$extractor}

    foreach($_name in $_obj.psobject.Properties.name){
      $main_obj | Add-Member -NotePropertyMembers @{$_name = $_obj.$_name} -Force
      #$main_obj | Add-Member -NotePropertyName $_name -NotePropertyValue $_obj.$_name -Force
    }
    $new_csv += $main_obj
  }
}


$integrated_field = $config.integrated_field
$integrated_obj = fn_Select_Sort_Unique $new_csv $integrated_field $extractor
$export_folder = (${HOME} + $config.export_folder)

$export_csv_path = ($export_folder + $config.export_csv_name)
fn_Export_CSV $integrated_obj $export_csv_path

$export_json_path = ($export_folder + $config.export_json_name)
fn_Export_JSON $integrated_obj $export_json_path

$integrated_obj | Format-Table

