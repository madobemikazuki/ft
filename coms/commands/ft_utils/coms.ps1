<#
  .\config\FT_Utils.json で指定した registed_source から
  申請会社名や雇用企業名を収集し一意化したJSONファイルを出力するスクリプト
  registed_source に記録されいる企業名には
  半角の(株)
  全角の（株）
  などが混在している。データベース側で企業名をバリデートしていない。
  このコマンドではその表記揺れを修正しない。
  オリジナルデータの真正さを棄損することはこのコマンドの本旨ではない。
#>
<# -------------------- ここから実行内容 --------------------- #>

. ..\ft_cores\FT_IO.ps1
. ..\ft_cores\FT_Array.ps1
. ..\ft_cores\FT_Object.ps1


$config_path = ".\config\FT_Utils.json"
$command_name = Split-Path -Leaf $PSCommandPath
[PSCustomObject]$config = ([FT_IO]::Read_JSON_Object($config_path)).$command_name

$registered_path = $config.paths.registed_source
[PSCustomObject[]]$registered_source = [FT_IO]::Read_JSON_Array((${Home} + $registered_path))
[PSCustomObject]$registered_obj = [FT_Array]::ToDict($registered_source, $config.primary_key);
[PSCustomObject]$script:coms_field = $config.field
[String[]]$script:company_class = $coms_field.psobject.Properties.name
#$company_class | Format-List

# Select
$new_dict = @{}
foreach ($_class in $company_class) {
  $new_dict[$_class] = [FT_Object]::Map([PSCustomObject]$registered_obj, $coms_field.$_class)
}

# Sort
$sorted_dict = @{}
foreach ($_class in $new_dict.keys) {
  $sorted_dict[$_class] = [FT_Array]::SortByUnique($new_dict.$_class, $coms_field.$_class[0])
}
Remove-Variable new_dict
#$sorted_dict | Format-Table

$flat_dict = @{}
foreach ($_class in $sorted_dict.keys) {
  $key_field = $coms_field.$_class[0]
  $value_field = $coms_field.$_class[1]
  $obj_arr = $sorted_dict.$_class
  $flat_dict[$_class] = [FT_Array]::Flat_KV($obj_arr, $key_field, $value_field)
}
Remove-Variable sorted_dict
#$flat_dict | Format-List

$completion_obj = [PSCustomObject]$flat_dict
Remove-Variable flat_dict

$export_path = (${HOME} + $config.paths.export_path)

if (Test-Path $export_path) {
  [PSCustomObject]$existing_Coms = [FT_IO]::Read_JSON_Object($export_path)
  $private:coms_class = $existing_Coms.psobject.properties.name
  
  $every = @()
  $private:appended_Coms = [PSCustomObject]@{}
  foreach ($_class in $coms_class) {
    $result = [FT_Object]::Compare_Count_Over($existing_Coms.$_class, $completion_obj.$_class)
    $every += $result
    #
    if ($result -eq $False) {
      continue
    }
  }
  # $every に $True が含まれていない なら何もせず終了する。
  # $every が全て$True であるなら反対の結果となる。
  # $True の反対であるか？=> True なのでこのif文が実行される
  # PowerShell に判別しやすいEvery メソッドは存在しない？
  if (!($False -in $every)) {
    $appended_Coms = [FT_Object]::Append_KV($existing_Coms, $completion_obj)
    Write-Host "増えよ 🌳🌳🌳"
    [FT_IO]::Write_JSON_Object($export_path, $appended_Coms)
    exit 0
  }
  Write-Host "本日は晴天なり 🌞"
  exit 0
}
else {
  Write-Host "生まれよ 🌳"
  [FT_IO]::Write_JSON_Object($export_path, $completion_obj)
  exit 0
}

