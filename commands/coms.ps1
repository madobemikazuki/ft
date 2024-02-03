<#
  申請会社や主管グループ等を収録するJSONを出力するスクリプト
#>
function private:fn_Read_JSON {
  Param(
    [Parameter(Mandatory = $True)]
    [ValidatePattern('\.json$')]$_path
  )
  $json = Get-Content -Path $_path -Encoding UTF8 | ConvertFrom-Json
  return $json
}

function private:fn_Read_CSV {
  Param(
    [Parameter(Mandatory = $True)]
    [ValidatePattern('\.csv$')]$_path
  )
  return Import-Csv -Path $_path -Encoding Default
}

function private:fn_Read {
  Param(
    [Parameter(Mandatory = $True)]
    [ValidatePattern("\.csv$|\.json$")]$_path
  )
  $_path
  switch -Regex ($_path) {
    "\.csv$" {
      return Import-Csv -Path $_path -Encoding Default
      break
    }
    "\.json$" {
      return Get-Content -Path $_path -Encoding UTF8 | ConvertFrom-Json
      break
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
    [Parameter(Mandatory = $True, Position = 1)][String[]][ref]$_keys
  )
  $key = $_keys[0]
  $_obj_list | Sort-Object -Property $key
}

function script:fn_Unique {
  Param(
    [Parameter(Mandatory = $True, Position = 0)][PSCustomObject[]][ref]$_obj_list,
    [Parameter(Mandatory = $True, Position = 1)][String[]][ref]$_keys
  )
  fn_Sort $_obj_list $_keys | Get-Unique -AsString
}

function private:fn_Extract_Set {
  Param(
    [Parameter(Mandatory = $True)]
    [PSCustomObject[]][ref]$psc_list,
    [Parameter(Mandatory = $True)]
    [String[]][ref]$keys
  )
  $private:coms = fn_Map $psc_list  $keys
  $private:coms_set = fn_Unique $coms $keys[0]
  return $coms_set
}


function private:fn_Create_KV {
  Param(
    [Parameter(Mandatory = $True)]
    [PSCustomObject[]][ref]$_psobj_list,
    [Parameter(Mandatory = $True)]
    [String][ref]$_key,
    [Parameter(Mandatory = $True)]
    [String][ref]$_value
  )
  #$private:kv = [PSCustomObject]@{}
  $private:new_KV = New-Object -TypeName psobject
  foreach ($_ in $_psobj_list) {
    Add-Member -InputObject $new_KV -MemberType NoteProperty -Name $_.$_key -Value $_.$_value
  }
  return $new_KV
}


function script:fn_Append_KV {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [PSCustomObject][ref]$_existing_KV,
    [Parameter(Mandatory = $True, Position = 1)]
    [PSCustomObject][ref]$_new_KV
  )  
  #上書きについて悩みたくないので、-Force で上書きを強制する。
  foreach ($key in $_new_KV.psobject.Properties.Name) {
    $_existing_KV | Add-Member -NotePropertyName $key -NotePropertyValue $_new_KV.$key -Force
  }
  # JOSNに不要な情報が格納されてしまうので、インデックスを指定している。
  return $_existing_KV[1]
}

function script:fn_Add_Objects{
  Param(
  [Parameter(Mandatory = $True, Position = 0)]
  [PSCustomObject[]][ref]$_exist,
  [Parameter(Mandatory = $True, Position = 1)]
  [PSCustomObject[]][ref]$_new,
  [Parameter(Mandatory = $True, Position = 2)]
  [String[]][ref]$_keys
  )
  $added = $_exist + $_new
  $new = $added | Sort-Object -Property $_keys[2] | Get-Unique -AsString
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




[PSCustomObject]$config = fn_Read_JSON ".\config\coms.json"

[PSCustomObject[]]$private:source = fn_Read_CSV (${home} + $config.paths.registed_source)
[String[]]$app_coms_field = $config.field.app_coms
[PSCustomObject[]]$app_coms_set = fn_Extract_Set ([ref]$source) ([ref]$app_coms_field)
[PSCustomObject]$app_coms_KV = fn_Create_KV ([ref]$app_coms_set) ([ref]$app_coms_field[0]) ([ref]$app_coms_field[1])


#雇用会社情報をマップする
$emp_coms_feild = $config.field.emp_coms
[PSCustomObject[]]$emp_coms = fn_Extract_Set ([ref]$source) ([ref]$emp_coms_feild)
$emp_coms_json_path = (${Home} + $config.paths.employ_coms)
#$emp_coms | Format-List

$app_coms_json_path = (${HOME} + $config.paths.application_coms)


switch (Test-Path $app_coms_json_path) {
  $True {
    # エクスポートパスが存在するなら 追記 する。
    [PSCustomObject]$existing_KV = fn_Read $app_coms_json_path
    $appended_KV = fn_Append_KV ([ref]$existing_KV) ([ref]$app_coms_KV)
    fn_Save $app_coms_json_path $appended_KV
 
    #TODO: 未実装 雇用会社情報の追記処理
    [PSCustomObject[]]$existing_emp_coms = (fn_Read $emp_coms_json_path)[1]
    $appended_emp_coms = fn_Add_Objects ([ref]$existing_emp_coms) ([ref]$emp_coms) ([ref]$emp_coms_feild)
    #$existing_emp_coms.Length
    #$appended_emp_coms.Length
    fn_Save $emp_coms_json_path $appended_emp_coms
    fn_Notifycation $config.command_name "出力完了🌲🌲🌲 : $app_coms_json_path"
    exit 0
  }
  Default {
    # エクスポートパスが存在しないなら 新規保存する。
    fn_SaveAs $app_coms_json_path $app_coms_KV
    fn_SaveAs $emp_coms_json_path $emp_coms
    fn_Notifycation $config.command_name "出力完了🌲 : $app_coms_json_path"
    exit 0
  }
}





