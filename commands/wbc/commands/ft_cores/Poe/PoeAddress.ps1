Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

class PoeAddress {

  # Single_Format と改名してもいいなぁ
  static [System.Collections.Generic.List[PoeObject]]Common_Format ([PSCustomObject]$_common_obj, [PSCustomObject]$_address_table) {
    $private:field = $_common_obj.PSObject.Properties.Name
    #共通情報を転記用のフォーマットに変換する
    $private:poe_obj_list = foreach ($_ in $field) {
      [PoeObject]::New($_,
        $_common_obj.$_,
        $_address_table.$_[0],
        $_address_table.$_[1])
    }
    return [System.Collections.Generic.List[PoeObject]]$poe_obj_list
  }

  # インクリメント処理しつつmapping処理をするので Unit_Format と似て非なる処理です。
  static [System.Collections.Generic.List[PoeObject]]List_Format ([PSCustomObject[]]$_obj_list, [PSCustomObject]$_address_table) {
    function private:fn_mapping {
      Param(
        [Parameter(Mandatory = $True, Position = 0)][PSCustomObject]$_obj,
        [Parameter(Mandatory = $True, Position = 1)][String[]]$_field,
        [Parameter(Mandatory = $True, Position = 2)][PSCustomObject]$_address,
        [Parameter(Mandatory = $True, Position = 3)][byte]$_index
      )
      foreach ($_ in $_field) {
        [PoeObject]::New($_, $_obj.$_, ($_address.$_[0] + $_index), $_address.$_[1])
      }
    }
    #申請者情報をExcel 転記用のフォーマットに変換する
    $private:field = $_address_table.PSObject.Properties.Name
    $poe_obj_list = foreach ($_obj in $_obj_list) {
      $index = $_obj_list.indexOf($_obj)
      fn_mapping $_obj $field $_address_table $index
    }
    return [System.Collections.Generic.List[PoeObject]]$poe_obj_list
  }


  static [System.Collections.Generic.List[PoeObject]]Unit_Format ([PSCustomObject[]]$_source_list, [PSCustomObject[]]$_address_table) {
    function private:fn_mapping {
      Param(
        [Parameter(Mandatory = $True, Position = 0)][PSCustomObject]$_obj,
        [Parameter(Mandatory = $True, Position = 1)][String[]]$_field,
        [Parameter(Mandatory = $True, Position = 2)][PSCustomObject]$_address
      )
      foreach ($_name in $_field) {
        [PoeObject]::New($_name, $_obj.$_name, $_address.$_name[0], $_address.$_name[1])
      }
    }
    $private:copied_list = $_source_list
    $private:printing_field = $_address_table[0].psobject.Properties.Name
    #申請者情報をExcel 転記用のフォーマットに変換する
    $formated_obj_list = foreach ($_obj in $copied_list) {
      $index = $copied_list.indexOf($_obj)
      fn_mapping $_obj $printing_field $_address_table[$index]
    }
    return [System.Collections.Generic.List[PoeObject]]$formated_obj_list
  }
  
}

