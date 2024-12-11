Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

class PoeAddress {

  # Single Format
  # 単ページに１セットだけ配置したいときのアドレステーブル
  # 日付、共通の組織名を記載するときに使用することもできる。
  <#
    _________________________________
    | aaa | bbbb | ccc | dddd | eee | 
    |                               |
    |                               |
    |                               |
    |_______________________________|
  #>
  static [System.Collections.Generic.List[PoeObject]]Single_Format ([PSCustomObject]$_common_obj, [PSCustomObject]$_address_table) {
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
  

  # List Format
  #　複数単位の情報セットを1ページに配置するときのアドレステーブル
  # アドレスをインクリメント処理しつつmapping処理をするので Unit_Format と似て非なる処理です。
  <#
  _________________________________
  | 0 | aaaa | bbbb | cccc | dddd | 
  | 1 | abbb | bccc | cddd | deee |
  | 2 | accc | bddd | ceee | dfff |
  |_______________________________|
  #>
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

  # Unit Format
  # 複数単位の情報セットを単ページに配置するときのアドレステーブル
  # 単位ごとにアドレステーブルを指定したいときに使用する。
  <#
  ________________________
  | 0 | a11|    | 1 | b11|
  | 0 | a22|    | 1 | b22|
  |______________________|
  | 2 | c11|    | 3 | d11|
  | 2 | c22|    | 3 | d22|
  |______________________|
  #>
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

