Param(
  [Parameter(Mandatory = $True, Position = 0)]
  [__ComObject]$_book,
  [Parameter(Mandatory = $True, Position = 1)]
  [PSCustomObject]$_printing
)

Set-StrictMode -Version 3.0

# 直前まで使用していたプリンタの情報取得
$default = Get-WmiObject Win32_Printer | Where-Object default

#今から使うプリンタを設定  プリンタ名が指定されないと例外が発生しスクリプトは止まる。
$printer = Get-WmiObject Win32_Printer | Where-Object name -eq $_printing.printer_name
$printer.SetDefaultPrinter()
Set-PrintConfiguration -PrinterName $printer.name -Color $_printing.color


# excel ファイル自体に両面印刷設定されていればOK
#Set-PrintConfiguration -PrinterName $_printing.printer_name -DuplexingMode TwoSidedShortEdge


$book = $_book
$start = [int16]$_printing.start_page
$end = [int16]$_printing.end_page
$copies = [int16]$_printing.number_of_copies

# プリントアウトする
$book.PrintOut.Invoke(@($start, $end, $copies))

#プリンタ設定をプリントアウト前の設定に戻す
$default.SetDefaultPrinter()
exit 0
