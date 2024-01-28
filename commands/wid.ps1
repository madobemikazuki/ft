
#. .\core\json\read_json.ps1
function private:fn_Read_JSON{
  Param(
    [Parameter(Mandatory = $True)]
    [ValidatePattern('\.json$')]$_path
  )
  $json  = Get-Content -Path $_path -Encoding UTF8| ConvertFrom-Json
  return $json
}

$private:config = fn_Read_JSON ".\config\wid_group.json"



try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $False
  $excel.DisplayAlerts = $False

  $book = $excel.Workbooks.Open(
    (${HOME} + $config.import.path),
    0,
    $true
  )
  $target_page = 1
  $Sheet = $book.Sheets.Item($target_page)

  #行 ( rows y軸) の設定 値取得対象の始点行から最終行の設定
  $starting_row = $config.import.starting_row
  $end_of_rows = $Sheet.UsedRange.Rows.Count + 1
  $select_rows_range = @($starting_row..$end_of_rows)
   
  # 列 ( columns x軸)の設定
  $starting_column = $config.import.starting_column
  $end_of_columns = $config.import.end_of_columns
  $columns = @($starting_column..$end_of_columns)
  $export_field = $config.export.field


  # PSCustomObject[]に格納する。
  [PScustomObject[]]$list = foreach ($_row in $select_rows_range) {
    # pscustomObject に格納する。
    # return object
    $obj = [PSCustomObject]@{}
    foreach ($_column in $columns) {
      $index = $columns.IndexOf($_column)
      $key = $export_field[$index]
      $value = $Sheet.Cells.Item($_row, $_column).Text
      $obj | Add-Member -MemberType NoteProperty -Name $key -Value $value
    }
    $obj
  }

  $new_list = $list |Sort-Object -Property '作業件名コード' -Descending
  $new_list | Format-Table


  # csv出力
  # csvに出力するには、エンコードはANSIでなければならない。
  # csvなんて使わないから コメントアウトしておく
  <#
  $csv_path = "${HOME}" + $config.export.csv_path
  New-Item -Path $csv_path -ItemType File -Force
  $new_list | Export-csv -path $csv_path -NoTypeInformation -Encoding Default
  #>

  # JSON出力 JSON出力はUTF8-bomでOK
  # JSONファイルをブラウザ上で読み込んだ場合、
  # 特定の文字列を検索するには F3 が有効である。
  $json_path = "${HOME}" + $config.export.json_path
  New-Item -Path $json_path -ItemType File -Force 
  $utf8_with_BOM = New-Object System.Text.UTF8Encoding $True
  [System.IO.File]::WriteAllLines($json_path, (ConvertTo-Json $new_list), $utf8_with_BOM)
}
catch [exception] {
  Write-Output "😢😢😢エラーをよく読んでね。"
  Write-Output $_
}
finally {
  $excel.Quit()
  @($excel, $book, $sheet) | ForEach-Object {
    if ($_ -ne $null) {
      [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($_)
    }
  }
}

# コマンド終了
exit 0
