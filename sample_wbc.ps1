try {
  $time =Measure-Command{
  $script:excel = New-Object -ComObject Excel.Application
  $excel.Visible = $False
  $excel.DisplayAlerts = $False
  }
  Write-Host ""
  Write-Host ""
  Write-Host $time.TotalSeconds.ToString("F2")"秒 : Excelの起動が完了するまでの経過時間"
  Write-Host ""
  Write-Host ""
  $source_path = "${HOME}\Downloads\from_T\登録\WBC受検用紙_登録_原紙.xlsx"
  
  $book = $excel.Workbooks.Open($source_path, 0, $true)
  
  $target_sheet_page = 1
  $sheet = $book.Worksheets.Item($target_sheet_page)
  $sheet.Cells.Item(5, 3) = "Hello my Scripts!"
  
  $output_path = "${HOME}\Downloads\output\登録\WBC受検用紙\success.xlsx"
  $book.SaveAs("$output_path")
  $book.Close()
  Write-Output "👍👍👍  出力先 : $output_path"
}
catch [exception] {
  Write-Output "😢😢😢エラーをよく読んでね。"
  $error[0].ToString()
  Write-Output $_
}
finally {
  @($sheet, $book) | ForEach-Object {
    if ($_ -ne $null) {
      [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($_)
    }
  }
  $excel.Quit()
}
  
exit 0