
class FT_IO {
  
  static [bool] Validate([String]$_path, [String]$_pattern) {
    $result = if ( (Test-Path $_path) -and [FT_IO]::Match_Extension($_path, $_pattern)) {
      $True
    }
    else { 
      Write-Host "Error :"$_path
      throw "上記 ファイルパス が存在しません。もしくは拡張子が不正です。"
    }
    return $result
  }


  static [bool] Match_Extension([String]$_path, [String]$_pattern) {
    return $_path -match $_pattern
  }


  static [PSCustomObject[]] Read_JSON_Array([String]$_path) {
    $private:json = if ([FT_IO]::Validate($_path, "\.json$")) {
      Get-Content -Path $_path -Encoding UTF8 | ConvertFrom-Json
    }
    return $json
  }

  static [PSCustomObject] Read_JSON_Object([String]$_path) {
    $private:json = if ([FT_IO]::Validate($_path, "\.json$")) {
      Get-Content -Path $_path -Encoding UTF8 | ConvertFrom-Json
    }
    return $json
  }

  static [void] Write_JSON_Array([String]$_path, [PSCustomObject[]]$_Object_List) {
    if (Test-Path $_path) {
      New-Item -Path $_path -ItemType File -Force
    }
    $utf8_with_BOM = New-Object System.Text.UTF8Encoding $True

    [System.IO.File]::WriteAllLines($_path, (ConvertTo-Json $_Object_List), $utf8_with_BOM)
    Write-Host '出力完了💩 ::'$_path
  }

  static [void] Write_JSON_Object([String]$_path, [PSCustomObject]$_Object) {
    if (!(Test-Path $_path)) {
      New-Item -Path $_path -ItemType File -Force
    }
    $utf8_with_BOM = New-Object System.Text.UTF8Encoding $True
        # ConvertTo-Json に -Depth を指定しないと深いインデントのオブジェクトが読み込まれないので注意。
    [System.IO.File]::WriteAllLines($_path, (ConvertTo-Json $_Object -Depth 3), $utf8_with_BOM)
    Write-Host '出力完了💩 ::'$_path
  }



  static[PSCustomObject[]] Read_CSV([String]$_path) {
    $private:csv = if ([FT_IO]::Validate($_path, "\.csv$")) {
      Import-Csv -Path $_path -Encoding Default
    }
    return $csv
  }

  static [void]Write_CSV() {

  }


  static [void] Write_CSV_UTF8([String]$_path, [PSCustomObject[]]$_csv_obj) {
    # ブラウザから参照するとき、エンコードが Default だと文字化けするのでUTF8指定する。
    # 強制上書きするよ。
    if ([FT_IO]::Match_Extension($_path, "\.csv$")) {
      $_csv_obj | Export-Csv -NotypeInformation -Path $_path -Delimiter "," -Encoding "UTF8" -Force
      Write-Host '出力完了::'$_path
    }
    else {
      Write-Host "ファイル拡張子が .csv ではないため保存できません。"
      throw
    }
  }


  static [PSCustomObject[]] Bind_As_CSV([Object[]]$_header, [Object[]]$_values, [String]$_delimiter) {
    [PSCustomObject[]]$obj_arr = $_values | ConvertFrom-Csv -Header $_header.Split($_delimiter)
    return $obj_arr
  }


  static [String[]] Find([String]$_Folder, [String]$_TargetName) {
    $private:file_path_list = @()
    $private:head = "*"
    $private:end = "*.*"
    $private:wild_card = @($head, $_TargetName, $end) -Join ""
    try {
      $file_path_list = (Get-ChildItem -Path $_Folder -File -Filter $wild_card).FullName
    }
    catch {
      Write-Host "エラー発生 :: $($_.Exception.Message)"
    }
    return $file_path_list
  }

  static [String[]] Read_ToArray([String]$_txt_path, [String]$_encode) {
    $array = if ([FT_IO]::Validate($_txt_path, "\.txt$")) {
      Get-Content -Path $_txt_path  -Encoding $_encode
    }
    return $array 
  }

  static [bool]Exists_Path([String]$_folder, [String[]]$_file_names) {
    return (Test-Path -Path $_folder -Include $_file_names)
  }

  static [String]Find_Latest_File([String]$_Folder, [String]$_TargetName) {
    $file_list = Get-ChildItem -Path $_Folder -File -Filter $_TargetName
    if ($null -eq $file_list) { throw "$_TargetName に該当するファイルが存在しません。" }
    $latest_file = ($file_list | Sort-Object LatestWriteTime -Descending)[0].FullName
    return $latest_file
  }

  static [void] Create_Empty_Files([String[]]$_file_path_arr) {
    foreach ($_ in @($_file_path_arr)) {
      if (!(Test-Path $_)) {
        # 空ファイルを作る
        New-Item -Path $_ -ItemType File -Force
      }
    }
  }

  static [void] Move_To_Waste([String[]]$_file_list, [String]$_waste_folder) {
    foreach ($_file in $_file_list) {
      Move-Item -Path $_file -Destination $_waste_folder
    }
  }

  static[void] Write_Log([String]$_path, [hashtable]$_log) {
    #TODO: もっとシンプルな実装ならいいかな。
    # *.ps1 | loop をスキップした. | object | 存在しない
    if (Test-Path $_path) {
      $log = $_log
      [PSCustomObject]$exists = [FT_IO]::Read_JSON_Object($_path)
      # PSCustomObject を HashTable に変換する
      foreach ($_key in $exists.psobject.Properties.name) {
        $log[$_key] = $exists.$_key
      }
      $ordered_key = $log.keys | Sort-Object -Descending
      $new_hash = [ordered]@{}
      foreach ($_key in $ordered_key) {
        $new_hash[$_key] = $log.$_key
      }
      [FT_IO]::Write_JSON_Object($_path, [PSCustomObject]$new_hash)
    }
    else { 
      [FT_IO]::Write_JSON_Object($_path, [PSCustomObject]$_log)
    }
  }
}

