Param(
  [Parameter(Mandatory = $True, Position = 0)]
  [ValidatePattern("^[0-9]{2}\-[0-9]{6}$")]
  [String[]]$_central_nums
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

. ..\ft_cores\FT_IO.ps1

$private:config = [FT_IO]::Read_JSON_Object(".\watchers\config\csvAPD.json")

$private:folder_path = (${HOME} + $config.folder)

try {
  $private:file_path = [FT_IO]::Find_Latest_File($folder_path, $config.file_name)
  $private:primary_key = $config.primary_key
  $private:extraction = $config.extraction_target
  
  $private:csv_obj = Import-Csv -Path $file_path -Encoding Default
  $private:exists = foreach ($_ in $csv_obj) {
    ($_.$primary_key -in $_central_nums)
  }
  
  if ($True -in $exists) {
    $private:target_obj = $csv_obj | Where-Object { ($_.$primary_key -in $_central_nums) }
    $private:now = (" : " + (Get-Date).ToString($config.datetime_format))
    Write-Host "最新情報" -NoNewline -BackgroundColor DarkGreen
    Write-Host $now

    # 関数内やスクリプトブロック内ではFormat-Table ではコンソールに表示されないため
    # Out-Host コマンドで強制的にコンソールへ表示する
    # PowerShell Format-Table 表示できない をキーワードとしてDuckDuckGoで検索し、
    # StackOverFlow のスレで確認した。
    $target_obj | Format-Table -Property $extraction | Out-Host
    #$target_obj | Select-Object -Property $extraction | Out-Host
  }
  else {
    Write-Host "該当者なし"
  }
}
catch {
  Write-Host "Erorr :" -NoNewline -BackgroundColor Red
  Write-Host $_.FullyQualifiedErrorId
}
finally{
  Remove-Variable config
}

