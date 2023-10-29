Param(
  [Parameter(Mandatory = $True, Position = 0)]
  [__ComObject]$_excel_book,
  [Parameter(Mandatory = $True, Position = 1)]
  [String]$_output_path
)

Set-StrictMode -Version 3.0

$_excel_book.SaveAs("$_output_path")
$_excel_book.Close()
Write-Output "👍👍👍  出力先 : $_output_path"

exit 0
