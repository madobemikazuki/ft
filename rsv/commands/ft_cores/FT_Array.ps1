class FT_Array {
  # PowerShell 5 では複数のプロパティを一度に指定することができません。

  static [PSCustomObject[]] Sort([PSCustomObject[]]$_arr, [String[]]$_keys) {
    # 昇順でソートします。
    $private:arr = $_arr
    $result = $arr | Sort-Object -Property $_keys
    return  $result
  }

  static [PSCustomObject[]] Map([PSCustomObject[]]$_arr, [String[]]$_keys){
    $private:arr = $_arr
    $result = $arr | Select-Object -Property $_keys
    return  $result
  }
  static [PSCustomObject[]]SortByUnique([PSCustomObject[]]$_arr, [String]$_primary_key){
    $private:arr = $_arr
    $result = $arr | Sort-Object -Property $_primary_key -Unique
    return  $result
  }

  static [PSCustomObject[]] SortDesc([PSCustomObject[]]$_arr, [String[]]$_keys) {
    # 降順でソートします。
    $private:arr = $_arr
    $result = foreach ($_key in $_keys) {
      $arr | Sort-Object -Property $_key -Descending
    }
    return  $result
  }

  static [PSCustomObject[]] Selection([PSCustomObject[]]$_arr, [String]$_key, [String]$_value) {
    $private:arr = $_arr
    $target_obj = foreach ($_obj in $arr) {
      if ($_obj.$_key -ne $_value) { continue }
      $_obj
    }
    return $target_obj
  }

  static [HashTable] ToDict([PSCustomObject[]]$_arr, [String]$_primary_key){
    $dict = @{}
    foreach($_ in $_arr){
      $dict[$_.$_primary_key] = $_ 
    }
    # hashtable key は文字列 Value はPSCustomObject
    return $dict
  }


  static [PSCustomObject[]] Null_Release([PSCustomObject[]]$_arr, [String]$_key) {
    $private:arr = $_arr
    $non_null_arr = foreach ($_ in $arr) {
      if ($null -eq $_.$_key) { continue }
      $_
    }
    return $non_null_arr
  }

  static [PSCustomObject[]] SearchObject([PSCustomObject[]]$_arr, [String]$_primary_key, [String[]]$_values){
    [PSCustomObject[]]$private:result = $_arr | Where-Object { $_.$_primary_key -in $_values } 
  return $result
  }
}

