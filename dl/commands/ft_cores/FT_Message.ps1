Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

class FT_Message{

  static execution ([String]$_message){
    Write-Host "実行:"$_message -BackgroundColor DarkBlue
  }
}

