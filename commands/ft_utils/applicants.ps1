Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

. ..\ft_cores\FT_Specific_Funcs.ps1
. ..\ft_cores\FT_Message.ps1
# ------------------------------------------------------ #
$private:config_path = ".\config\FT_Utils.json"
$private:command_name = Split-Path -Leaf $PSCommandPath
$private:config = fn_SF_Read_Config $config_path $command_name
[FT_Message]::execution($command_name)

$private:tasks = ($config.tasks).psobject.Properties.name

foreach ($_task in $tasks) {

  $t_config = ($config.tasks).$_task
  $odd_field = "ed_dict"
  $Odd = fn_SF_Read_Config ((${Home} + $t_config.odd_dict_path)) $odd_field
  #$Odd.GetType()
  
  #申請予定者情報
  $private:cndd_path = $t_config.candidates_Path
  $private:cndd_lookup_key = $t_config.candidates_primary_key
  $private:candidates_dict = fn_SF_Read_as_Dict $cndd_path $cndd_lookup_key
  #$candidates_keys

  #予約者情報
  $private:rsv_path = $t_config.reserved_Path
  $private:rsv_lookup_key = $t_config.reserved_primary_key
  $private:reserved_dict = fn_SF_Read_as_Dict $rsv_path $rsv_lookup_key
  #$reserved_dict.Values | Format-List

  # $_task ごとに抽出、整形したjsonファイルを出力する
  # 放管手帳継続発行後はうまくいかないかもしれない。
  switch ($_task) {
    'r' { # 登録モード
      $marged_arr = fn_SF_Marge_Dicts $candidates_dict $reserved_dict
      #$marged_arr | Format-List
      $applicants_arr = fn_SF_Registration_Format $marged_arr $Odd
      fn_SF_Write_JSON_Array ($t_config.export_Path) $applicants_arr
      break
    }
    'c' { #解除モード
      $marged_arr = fn_SF_Marge_Dicts $candidates_dict $reserved_dict
      $applicants_arr = fn_SF_Cancellation_Format $marged_arr $Odd
      fn_SF_Write_JSON_Array ($t_config.export_Path) $applicants_arr
      break
    }
  }
}
