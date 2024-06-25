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
      [PSCustomObject[]]$private:obj = Import-Csv -Path $_path -Encoding Default
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
  $private:kv = [PSCustomObject]@{}
  $private:key = $_key
  $private:value = $_value
  foreach ($_ in $_psobj_list) {
    Add-Member -InputObject $kv -NotePropertyName $_.$key -NotePropertyValue $_.$value -Force
  }
  return $kv
}


function script:fn_Append_KV {
  Param(
    [Parameter(Mandatory = $True, Position = 0)]
    [PSCustomObject][ref]$_existing_KV,
    [Parameter(Mandatory = $True, Position = 1)]
    [PSCustomObject][ref]$_new_obj
  )
  [PSCustomObject]$private:exist = $_existing_KV
  [String[]]$private:exist_field = $exist.psobject.Properties.Name
  [String[]]$private:new_obj_field = $_new_obj.psobject.Properties.Name
  
  foreach ($_key in $new_obj_field) {
    if($_key -in $exist_field){ continue }
    Add-Member -InputObject $exist -NotePropertyMembers @{ $_key = $_new_obj.$_key }
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


[PSCustomObject]$script:config = fn_Read ".\config\coms.json"
[PSCustomObject[]]$script:regists_source = fn_Read (${Home} + $config.paths.registed_source)

[PSCustomObject]$script:coms_field = $config.field
[String[]]$script:company_styles = $coms_field.psobject.Properties.name

$script:integrated_obj = [PSCustomObject]@{}

foreach ($_style_key in $company_styles) {
  $coms_set = fn_Extract_Set ([ref]$regists_source) ([ref]$coms_field.$_style_key)
  $kv_coms = fn_To_KV ([ref]$coms_set) ([ref]$coms_field.$_style_key[0]) ([ref]$coms_field.$_style_key[1])
  #$kv_coms | Format-List
  $integrated_obj | Add-Member -NotePropertyMembers @{ $_style_key = $kv_coms }
}

#$integrat_obj.gettype()
#$integrat_obj | Format-List

$export_path = (${HOME} + $config.paths.export_path)


if (Test-Path $export_path) {
  Write-Host "増えよ"
  [PSCustomObject]$existing_KV = fn_Read $export_path
  #$existing_KV | Format-List
  $appended_KV = fn_Append_KV ([ref]$existing_KV) ([ref]$integrated_obj)
  fn_Save $export_path $appended_KV
  fn_Notifycation $config.command_name "追記完了🌲🌲🌲 : $export_path"
  exit 0
}
else {
  Write-Host "生まれよ"
  fn_SaveAs $export_path $integrated_obj
  fn_Notifycation $config.command_name "新規出力完了🌲 : $export_path"
  exit 0
}
