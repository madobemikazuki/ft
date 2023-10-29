Set-StrictMode -Version 3.0

#New-Object -ComObject でCOMオブジェクトを使用。
$excel = New-Object -ComObject Excel.Application


function read_xlsx_nonVisible {
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory = $true)]
    [String]$temp_path
  )

  try {
    #Get-ChildItem と .FullName でファイルの絶対パスを取得。
    $xlsx_path = (Get-ChildItem $temp_path).FullName

    #.Visible = $false でExcelを表示しないで処理を実行できる。
    $excel.Visible = $False

    # 上書き保存時に表示されるアラートなどを非表示にする
    $excel.DisplayAlerts = $False

    #.Workbooks.Open(ファイル名, リンクの更新方法, 読み取り専用) でExcelを開きます1。
    # リンクの更新方法が 0 の場合は何もしない。
    return $excel.Workbooks.Open($xlsx_path, 0, $true)
  }
  catch  [exception] {
    <#Do this if a terminating exception happens#>
    $error[0].ToString()
  }
}

function read_xlsx_Visible {
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory = $true)]
    [String]$temp_path
  )

  try {
    #Get-ChildItem と .FullName でファイルの絶対パスを取得。
    $xlsx_path = (Get-ChildItem $temp_path).FullName

    #.Visible = $false でExcelを表示しないで処理を実行できる。
    $excel.Visible = $True

    # 上書き保存時に表示されるアラートなどを非表示にする
    $excel.DisplayAlerts = $False

    #.Workbooks.Open(ファイル名, リンクの更新方法, 読み取り専用) でExcelを開きます1。
    # リンクの更新方法が 0 の場合は何もしない。
    return $excel.Workbooks.Open($xlsx_path, 0, $true)
  }
  catch  [exception] {
    <#Do this if a terminating exception happens#>
    $error[0].ToString()
  }
}
function quit_excel {
  $excel.Quit()
  [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
}
