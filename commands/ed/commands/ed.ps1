Param(
  [Parameter(Mandatory = $True, Position = 0)]
  [ValidatePattern("ts|tp|ih|TS|TP|IH")]
  [String]$_Tgroup,
  [Parameter(Mandatory = $True, Position = 1)]
  [ValidatePattern("c|d|cd|j|C|D|CD|J")]
  [String]$_Kinds,
  [Parameter(Mandatory = $true, Position = 2)]
  [ValidatePattern("^\d{2}\b\-\b\d{6}$")]
  [String[]]$_central_nums
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"


<#
1.事前申請者の登録予約情報から中央登録番号の該当者を検索する
2.既存の登録者の中から中登番号の該当者を検索する
上記、いずれかの該当者の情報を参照し、教育実施記録を出力する。
#>



# 初期化処理 --------------------------------------------

. .\ft_cores\FT_IO.ps1
. .\ft_cores\FT_Name.ps1
. .\ft_cores\FT_Array.ps1
. .\ft_cores\FT_Dict.ps1
. .\ft_cores\Poe\PoeObject.ps1
. .\ft_cores\Poe\PoeAddress.ps1

[PSCustomObject]$config = [FT_IO]::Read_JSON_Object(".\config\ed.json")
#$config.TS.export_folder 
#Test-Path (${HOME} + $config.gZEN_csv)


$reserved_source_Path = (${HOME} + $config.reserved_info.path)
[PSCustomObject[]]$reserved_arr = [FT_IO]::Read_JSON_Array($reserved_source_Path)

$primary_key = $config.primary_key
$reserved_dict = [FT_Array]::ToDict($reserved_arr, $primary_key)


$header = $config.extraction
# 申請用オブジェクト生成 --------------------------------------


# 登録予約者情報に該当者が存在する場合 not equals
if ([FT_Dict]::Every($reserved_dict, $_central_nums)) {
  $targets = [FT_Dict]::Search($reserved_dict, $_central_nums)
  $script:applicants_dict = [FT_Dict]::Selective($targets, $header)
  Write-Host "登録予約者のなかにおったよ。"
  Remove-Variable reserved_source_Path, reserved_arr, targets
}

# 登録予約者に該当者が存在しない場合 既存の登録者の中から探す
if (![FT_Dict]::Every($reserved_dict, $_central_nums)) {
  Write-Host "予約情報のなかにはおらんやったよ。"

  $private:registerers_path = (${Home} + $config.registerers_info.path)
  $private:registerers_arr = [FT_IO]::Read_JSON_Array($registerers_path)
  $private:registeres_dict = [FT_Array]::ToDict($registerers_arr, $primary_key)
  if ([FT_Dict]::Every($registeres_dict, $_central_nums)) {
    $private:targets = [FT_Dict]::Search($registeres_dict, $_central_nums)
    $script:applicants_dict = [FT_Dict]::Selective($targets, $header)
    Write-Host "既登録者のなかにおったよ。"
    #$applicants_dict.Values | Format-List
  }
  else {
    Write-Host "既登録者のなかにもおらんかったよ。"
    exit 404
  }
  #無限ループでアプリを使い回すなら、変数を消すと速度低下になるか？
  Remove-Variable primary_key, registerers_path, registerers_arr, registeres_dict
}

# まさか登録予定者と定期受検者を混在させて入力することはないだろう。
#$applicants_dict.Values | Format-Table


# エクセルオブジェクトへマッピング、書き出し 副作用満載-------------
try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $False
  $excel.DisplayAlerts = $False
  #.Workbooks.Open(ファイル名, リンクの更新方法, 読み取り専用) でExcelを開きます1。
  # リンクの更新方法が 0 の場合は何もしない。
  $book = $excel.Workbooks.Open(
    (${HOME} + $config.template_path),
    0,
    $true
  )

  foreach ($_applicant in $applicants_dict.Values) {

    foreach ($_kind in $_Kinds.ToCharArray()) {
      $page = $config.$_Tgroup.sheet_pages.$_kind
      $formatted_obj = [PoeAddress]::Common_Format($_applicant, $config.$_kind.address_table)
      $formatted_obj[0].value = ($config.$_kind.sandwitch) -replace @($config.replacement, $formatted_obj[0].value)
      #$formatted_obj | Format-Table
      
      $sheet = $book.Worksheets.Item($page)
      foreach ($_obj in $formatted_obj) {
        $sheet.Cells.Item($_obj.point_x, $_obj.point_y) = $_obj.value
      }
      #プリントアウトする
      if ($config.printable) {
        $book.PrintOut.Invoke(@($page, $page, [int16]$config.printing.number_of_copies))
      }
    }

    # exportする
    $export_name = @(
      "${HOME}",
      $config.$_Tgroup.export_folder,
      $config.File.head_name,
      $_applicant.($config.File.applicant),
      $config.File.extension
    ) -join ""
    $export_path = [FT_Name]::One_Liner($export_name)

    # 空ファイルを作成
    New-Item -Path $export_path -ItemType File -Force
    
    # 空ファイルに書き込む
    $book.SaveAs($export_path)
    Write-Output "🌸🌸🌸  出力先 : $export_path" 
  }
  $excel.Quit()
}
catch [exception] {
  Write-Output "😢😢😢エラーをよく読んでね。"
  $error[0].ToString()
  Write-Output $_
  $excel.Quit()
}
finally {
  $excel.Quit()
  $excel = $null
  [System.GC]::Collect()
  foreach ($_ in @( $sheet, $book , $excel)) {
    if ($null -ne $_) {
      [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($_)
    }
  }
  [System.GC]::Collect()
}
exit 0

