<#
  申請会社や主管グループ等を収録するJSONを出力するスクリプト
#>

function private:fn_Read {
  Param(
    [Parameter(Mandatory = $True)]
    [ValidatePattern("\.csv$|\.json$")]$_path
  )
  #$_path
  switch -Regex ($_path) {
    "\.csv$" {
      [PSCustomObject[]]$obj = Import-Csv -Path $_path -Encoding Default
      return $obj
    }
    "\.json$" {
      $obj = Get-Content -Path $_path -Encoding UTF8 | ConvertFrom-Json
      return $obj
    }
    Default {
      Write-Host "拡張子が該当しないので終了。"
      exit 0
    }
  }
}

function script:fn_Map {
  Param(
    [Parameter(Mandatory = $True, Position = 0)][PSCustomObject[]][ref]$_obj_list,
    [Parameter(Mandatory = $True, Position = 1)][String[]][ref]$_keys
  )
  $_obj_list | Select-Object -Property $_keys
}

function script:fn_Sort {
  Param(
    [Parameter(Mandatory = $True, Position = 0)][PSCustomObject[]][ref]$_obj_list,
    [Parameter(Mandatory = $True, Position = 1)][String][ref]$_keys
  )
  $_obj_list | Sort-Object -Property $_keys
}

function script:fn_Unique {
  Param(
    [Parameter(Mandatory = $True, Position = 0)][PSCustomObject[]][ref]$_obj_list,
    [Parameter(Mandatory = $True, Position = 1)][String[]][ref]$_keys
  )
  fn_Sort $_obj_list $_keys[0] | Get-Unique -AsString
}

function private:fn_Extract_Set {
  Param(
    [Parameter(Mandatory = $True)]
    [PSCustomObject[]][ref]$psc_list,
    [Parameter(Mandatory = $True)]
    [String[]][ref]$keys
  )
  $private:coms = fn_Map $psc_list $keys
  $private:coms_set = fn_Unique $coms $keys[0]
  return $coms_set
}


function private:fn_To_KV {
  Param(
    [Parameter(Mandatory = $True)]
    [PSCustomObject[]][ref]$_psobj_list,
    [Parameter(Mandatory = $True)]
    [String][ref]$_key,
    [Parameter(Mandatory = $True)]
    [String][ref]$_value
  )
  #配列を単一のPSCustomObjectへ変換し返す => {"T000":"ahaha.com", "T001":"Ohoho.com"}
  $_kv = [PSCustomObject]@{}
  foreach ($_obj in $_psobj_list) {
    Add-Member -InputObject $_kv -NotePropertyName $_obj.$_key -NotePropertyValue $_obj.$_value -Force
  }
  return $_kv
}


function script:fn_Append_KV {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [PSCustomObject][ref]$_existing_KV,
    [Parameter(Mandatory = $True, Position = 1)]
    [PSCustomObject][ref]$_new_obj
  )
  [PSCustomObject]$private:exist = $_existing_KV
  foreach ($_key in $_new_obj.psobject.Properties.Name) {
    Add-Member -InputObject $exist -NotePropertyName $_key -NotePropertyValue $_new_obj.$_key -Force
  }
  # JOSNに不要な情報が格納されてしまうので、インデックスを指定している。
  #return $_existing_KV[0]
  return $exist
}

function script:fn_Add_Objects {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [PSCustomObject[]][ref]$_exist,
    [Parameter(Mandatory = $True, Position = 1)]
    [PSCustomObject[]][ref]$_new,
    [Parameter(Mandatory = $True, Position = 2)]
    [String[]][ref]$_key
  )
  $added = $_exist + $_new
  [PSCustomObject[]]$new = $added | Sort-Object -Property $_key | Get-Unique -AsString
  return $new
}


function script:fn_Notifycation {
  Param(
    [String]$title,
    [String]$message
  )
  Add-Type -AssemblyName System.Windows.Forms
  $MUTEX_NAME = "Global\mutex" #多重起動チェック用

  try {
    $mutex = New-Object System.Threading.Mutex($False, $MUTEX_NAME)
    #多重起動チェック
    if ($mutex.WaitOne(0, $False)) {
      $notify_icon = New-Object Windows.Forms.NotifyIcon

      $notify_icon.Icon = [Drawing.SystemIcons]::Application
      $notify_icon.Visible = $True

      $notify_icon.BalloonTipText = "$title :  $message"
      $notify_icon.ShowBalloonTip(1)

      # $_second 秒待機して通知を非表示にする。
      $notify_icon.Visible = $False
    }
  }
  finally {
    $notify_icon.Dispose()
    $mutex.ReleaseMutex()
    $mutex.Close()
    $mutex.Dispose()
    exit
  }
}



function script:fn_Save {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [String]$_path,
    [Parameter(Mandatory = $True, Position = 1)]$_object
  )
  $utf8_with_BOM = New-Object System.Text.UTF8Encoding $True
  [System.IO.File]::WriteAllLines($_path, (ConvertTo-Json $_object), $utf8_with_BOM)
}

function script:fn_SaveAs {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [String]$_path,
    [Parameter(Mandatory = $True, Position = 1)]$_object
  )
  New-Item -Path $_path -ItemType File -Force
  fn_Save $_path $_object
}



<# -------------------- ここから実行内容 --------------------- #>


[PSCustomObject]$config = fn_Read ".\config\coms.json"
[PSCustomObject[]]$script:regists_source = fn_Read (${Home} + $config.paths.registed_source)
[String[]]$app_coms_field = $config.field.app_coms
$app_coms_field
#全ての登録者から申請会社情報のみ抽出する
[PSCustomObject[]]$app_coms_set = fn_Extract_Set ([ref]$regists_source) ([ref]$app_coms_field)
#Write-Host "app_coms_set"
#$app_coms_set | Format-Table

[PSCustomObject]$kv_app_coms = fn_To_KV $app_coms_set $app_coms_field[0] $app_coms_field[1]
Write-Host "kv_app_coms"
$kv_app_coms | Format-List


#雇用会社情報をマップする
$emp_coms_feild = $config.field.emp_coms
[PSCustomObject[]]$script:emp_coms_set = fn_Extract_Set ([ref]$regists_source) ([ref]$emp_coms_feild)
$emp_coms_json_path = (${Home} + $config.paths.employ_coms)
#$emp_coms | Format-List

$app_coms_json_path = (${HOME} + $config.paths.application_coms)


switch (Test-Path $app_coms_json_path) {
  $True {
    # エクスポートパスが存在するなら 追記 する。
    [PSCustomObject]$existing_KV = fn_Read $app_coms_json_path
    $existing_KV | Format-List
    $appended_KV = fn_Append_KV ([ref]$existing_KV) ([ref]$kv_app_coms)
    fn_Save $app_coms_json_path $appended_KV
 
    # 雇用会社情報の追記処理
    [PSCustomObject[]]$existing_emp_coms = (fn_Read $emp_coms_json_path)[1]
    $appended_emp_coms = fn_Add_Objects ([ref]$existing_emp_coms) ([ref]$emp_coms_set) ([ref]$emp_coms_feild[2])
    fn_Save $emp_coms_json_path $appended_emp_coms
    fn_Notifycation $config.command_name "出力完了🌲🌲🌲 : $app_coms_json_path"
    exit 0
  }
  Default {
    # エクスポートパスが存在しないなら 新規保存する。
    fn_SaveAs $app_coms_json_path $kv_app_coms
    fn_SaveAs $emp_coms_json_path $emp_coms_set
    fn_Notifycation $config.command_name "出力完了🌲 : $app_coms_json_path"
    exit 0
  }
}

