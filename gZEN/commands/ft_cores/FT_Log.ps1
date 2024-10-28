class FT_Log {
  
  static $extension_pattern = "\.json$"

  static [HashTable]Create($_command_name, $log_type, $_target, $_result, $_purpose) {
    $primary_key = (Get-Date).ToString()
    $hash = [ordered]@{}
    $hash[$primary_key] = @{
      command_name = $_command_name;
      log_type     = $log_type;
      target       = $_target;
      result       = $_result;
      purpose      = $_purpose;
    }
    return $hash
  }

  # $_path が存在し、かつ拡張子が $_pattern にマッチすること
  static [bool] Validate ([String]$_path, [String]$_pattern) {
    $result = if ( (Test-Path $_path) -and [FT_Log]::Match_Extension($_path, $_pattern)) {
      $True
    }
    else { 
      Write-Host "Error :"$_path
      throw "上記 ファイルパス が存在しません。もしくは拡張子が不正です。"
    }
    return $result
  }

  # $_path が $_pattern にマッチすること
  static [bool] Match_Extension([String]$_path, [String]$_pattern) {
    return $_path -match $_pattern
  }

  static [PSCustomObject] Read_JSON([String]$_path) {
    $private:json = if ([FT_Log]::Validate($_path, [FT_Log]::extension_pattern)) {
      Get-Content -Path $_path -Encoding UTF8 | ConvertFrom-Json
    }
    return $json
  }

  static [void] Write_JSON([String]$_path, [PSCustomObject]$_Object) {
    if (!(Test-Path $_path)) {
      New-Item -Path $_path -ItemType File -Force
    }
    $utf8_with_BOM = New-Object System.Text.UTF8Encoding $True
    # ConvertTo-Json に -Depth を指定しないと深いインデントのオブジェクトが読み込まれないので注意。
    [System.IO.File]::WriteAllLines($_path, (ConvertTo-Json $_Object -Depth 3), $utf8_with_BOM)
  }

  static[void] Write([String]$_path, [hashtable]$_log) {
    #TODO: もっとシンプルな実装ならいいかな。
    # *.ps1 | loop をスキップした. | object | 存在しない
    if ([FT_Log]::Validate($_path, [FT_Log]::extension_pattern)) {
      $log = $_log
      [PSCustomObject]$exists = [FT_Log]::Read_JSON($_path)
      # PSCustomObject を HashTable に変換する
      foreach ($_key in $exists.psobject.Properties.name) {
        $log[$_key] = $exists.$_key
      }
      $ordered_key = $log.keys | Sort-Object -Descending
      $new_hash = [ordered]@{}
      foreach ($_key in $ordered_key) {
        $new_hash[$_key] = $log.$_key
      }
      [FT_Log]::Write_JSON($_path, [PSCustomObject]$new_hash)
    }
    else {
      if ([FT_Log]::Match_Extension($_path, [FT_Log]::extension_pattern)) {
        [FT_Log]::Write_JSON($_path, [PSCustomObject]$_log)
      }
    }
  }
}

