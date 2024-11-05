class FT_Object{

  # @{ name : {key:value} } な PSCustomObject を対象としたMapping処理
  static [PSCustomObject]Map([PSCustomObject]$_obj, [String[]]$_field){
    $private:result = foreach ($primary_key in $_obj.psobject.properties.name) {
      $_obj.$primary_key | Select-Object -Property $_field
    }
    return $result
  }


  static [PSCustomObject]Append_KV([PSCustomObject]$_existing_KV, [PSCustomObject]$_new_obj){
    [PSCustomObject]$private:exist = $_existing_KV
    [String[]]$private:exist_field = $exist.psobject.Properties.Name
    [String[]]$private:new_obj_field = $_new_obj.psobject.Properties.Name
    
    foreach ($_key in $new_obj_field) {
      if ($_key -in $exist_field) { continue }
      Add-Member -InputObject $exist -NotePropertyMembers @{ $_key = $_new_obj.$_key }
    }
    return $exist
  }

  # 二つのオブジェクトのプロパティネームのcount数がdef_objの方が大きいか比較する。
  static [Bool]Compare_Count_Over([PSCustomObject]$_ref_obj, [PSCustomObject]$_def_obj){
    $r = $_ref_obj.PSObject.Properties.name.count
    $d = $_def_obj.PSObject.Properties.name.count
    #Write-Host $r $d
    return ($d -gt $r)
  }
}
