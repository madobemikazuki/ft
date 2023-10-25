Set-StrictMode -Version 3.0

function write_xslx {
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory = $true)]
    [__ComObject]$book,
    [String]$output_path
  )
  Write-Output "👍👍👍  出力先 : $output_path"
  $book.SaveAs("$output_path")
  $book.Close()
}