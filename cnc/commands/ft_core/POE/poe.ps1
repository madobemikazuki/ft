Param(
  [Parameter(Mandatory = $True, Position = 0)]
  [PSCustomObject]$_poe_config,
  [Parameter(Mandatory = $True, Position = 1)]
  [PSCustomObject[]]$_info_obj_list
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"


function fn_GenerateFilePath {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [String[]]$_names,
    [Parameter(Mandatory = $True, Position = 1)]
    [PSCustomObject]$_export_config
  )  
  $private:folder = (${HOME} + $_export_config.folder)
  $private:head_name = $_export_config.file_name.first
  $private:names = $_names -join $_export_config.file_name.conjunction
  $private:shorten_names = $names -replace $_export_config.file_name.replaces
  $private:extension = $_export_config.file_name.extension
  return ($folder + $head_name + $shorten_names + $extension)
}


[byte]$script:info_length = $_info_obj_list.Length

$script:address_table = $_poe_config.address_table
$script:poe_field = $_poe_config.printing.printig_field
$script:name_property = "漢字氏名"
$script:export_config = $_poe_config.export

# ユニット処理とチャンク処理
if ($_poe_config.printing.style -eq "chunk") {
  $chunk_size = [int16]($_poe_config.address_table.Length)
  # $chunk_size
  if ($chunk_size -lt $info_length) {
    Write-Host "Chunk_Processing : チャンク転記処理するよ"
    $chunked_arr = . .\core\chunked_array.ps1 $_info_obj_list $chunk_size
    Write-Host "チャンク配列の数 : " $chunked_arr.Length

    foreach ($_chunk in $chunked_arr) {
      $names = foreach ($_ in $_chunk) { $_.$name_property }
      $export_path = fn_GenerateFilePath $names $export_config
      $formated_list = . .\core\posting_format.ps1 $_chunk $poe_field $address_table
      . .\core\transcription.ps1 $formated_list $_poe_config $export_path
    }
  }
  # 
  # 0 < $info_length && $info_length < ($chunk_size + 1)
  if ((0 -lt $info_length) -and ($info_length -lt ($chunk_size + 1))) {
    Write-Host "Unit_Processing : ユニット転記処理するよ"
    $names = foreach ($_ in $_info_obj_list) { $_.$name_property }
    $export_path = fn_GenerateFilePath $names $export_config
    $formated_list = . .\core\posting_format.ps1 $_info_obj_list $poe_field $address_table
    . .\core\transcription.ps1 $formated_list $_poe_config $export_path
  }
  exit 0
}


if ($_poe_config.printing.style -eq "list") {
  Write-Host $info_length
  if (0 -lt $info_length) {
    Write-Host "List_Processing"
    Write-Host "リスト転記処理するよ"
    $names = foreach ($_ in $_info_obj_list) { 
      if ($name_property -eq $_.Name) { $_.Value }
    }
    $export_path = fn_GenerateFilePath $names $export_config
    . .\ft_core\POE\core\transcription.ps1 $_info_obj_list $_poe_config $export_path
  }
  exit 0
}


#TODO: しばらく実装する必要がない。
if ($_poe_config.printing.style -eq "single") {
  Write-Host "Single_Processing"
  Write-Host "シングル転記処理するよ"
  exit 0
}

exit 0

