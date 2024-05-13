Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

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

# 事前申請情報をJSONへ変換変換
$gZEN_config = fn_Read ".\config\gzen.json"
$gZEN_target = $gZEN_config.gZEN_targets
$PAN_folder = (${Home} + $gZEN_config.temp_folder)
$files = (Get-ChildItem -Path $PAN_folder -File -Filter $gZEN_target).fullname
if($files.length -gt 0){
  . .\coms.ps1
  . .\gZEN.ps1
}

# 予約済み情報の生成 => *.json 出力
$regists_path = (${Home} + "\Downloads\TEMP\登録者管理リスト_coh.csv")
if(Test-Path -Path $regists_path){
  . .\rsv.ps1
  . .\bind_c.ps1
  . .\bind_r.ps1
}

