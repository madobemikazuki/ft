class FT_Array {
  # PowerShell 5 では複数のプロパティを一度に指定することができません。

  static [PSCustomObject[]] Sort([PSCustomObject[]]$_arr, [String[]]$_keys) {
    # 昇順でソートします。
    $private:arr = $_arr
    $result = $arr | Sort-Object -Property $_keys
    return  $result
  }

  static [PSCustomObject[]] Map([PSCustomObject[]]$_arr, [String[]]$_keys) {
    $private:arr = $_arr
    $private:result = $arr | Select-Object -Property $_keys
    return  $result
  }

  static [String[]] V([PSCustomObject[]]$_arr, [String]$_key) {
    $private:arr = $_arr
    $private:result = foreach ($_ in $arr) {
      $_.$_key
    }
    return $result
  }

  static [PSCustomObject[]]SortByUnique([PSCustomObject[]]$_arr, [String]$_primary_key) {
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

  static [HashTable] ToDict([PSCustomObject[]]$_arr, [String]$_primary_key) {
    $dict = @{}
    foreach ($_ in $_arr) {
      $dict[$_.$_primary_key] = $_ 
    }
    # hashtable key は文字列 Value はPSCustomObject
    return $dict
  }


  # 要素が Null ではないものを返す
  static [PSCustomObject[]] Null_Release([PSCustomObject[]]$_arr, [String]$_key) {
    $private:arr = $_arr
    $non_null_arr = foreach ($_ in $arr) {
      if ($null -eq $_.$_key) { continue }
      $_
    }
    return $non_null_arr
  }

  static [PSCustomObject[]] SearchObject([PSCustomObject[]]$_arr, [String]$_primary_key, [String[]]$_values) {
    [PSCustomObject[]]$private:result = $_arr | Where-Object { $_.$_primary_key -in $_values } 
    return $result
  }

  static [PSCustomObject]Flat_KV([PSCustomObject[]]$_arr, [String]$_k, [String]$_v) {
    $private:k = $_k
    $private:v = $_v
    $private:KV = [PSCustomObject]@{}
    foreach ($_obj in $_arr) {
      if ([String]::IsNullOrEmpty($_obj.$k) -or [String]::IsNullOrEmpty($_obj.$v)) { continue }
      Add-Member -InputObject $KV -NotePropertyName $_obj.$k -NotePropertyValue $_obj.$v -Force
    }
    return $KV
  }

  static [PSCustomObject[]]Jugged([PSCustomObject[]]$_arr, [int16]$_range) {
    $private:new_arr = [PSCustomObject[]]@()
    if ($_range -lt $_arr.Length) {    
      $private:part = [Math]::Floor($_arr.length / $_range)
      #Write-Host $part " : 除算した数(少数点以下を切捨て)" 
      $private:remainder = $_arr.length % $_range
      #Write-Host $remainder " : 除算して余った数"
      
      $start_index = 0
      foreach ($_ in @(1..$part)) {
        $x = ($start_index + $_range - 1)
        [PSCustomObject[]]$split_Item = $_arr[$start_index..$x]
        $start_index += $_range
        $new_arr += , $split_Item
      }

      if ($remainder -gt 0) {
        # $remainder がゼロより大きければ端数の配列をジャグ配列として返す
        $new_arr += , [PSCustomObject[]]$_arr[-1.. - $remainder]
      }
    }
    return $new_arr
  }

  #TODO: 追加したから検証しろ
  # ------------------------------------------------------------------------------------
  # 重複しない要素の配列を返す
  static [PSCustomObject[]]No_Duplicates ([PSCustomObject[]]$_exists_arr, [PSCustomObject[]]$_addition_arr, [String]$_key) {
    # obj_list に $_value_list を照会して重複しない要素の配列を返す
    $private:value_list = foreach ($_ in $_addition_arr) { $_.$_key }
    $private:result = $_exists_arr | Where-Object { $value_list -notcontains $_.$_key }
    return $result
  }

  #追加リスト($_addition_list)の要素に不完全リストのプロパティが含まれていること
  # かつ、当該プロパティの値に $null や '' が含まれていない要素を配列にして返す
  static [PSCustomObject[]]Filled([PSCustomObject[]]$_addition_arr, [PSCustomObject[]]$_incomplete_arr, [String]$_target_key) {
    $private:incomplete_values = [FT_Array]::V($_incomplete_arr, $_target_key)

    [PSCustomObject[]]$private:filled_arr = $_addition_arr | Where-Object {
      #追加リスト($_addition_list)の要素に不完全リストのプロパティが含まれていること
      ($incomplete_values.Contains($_.$_target_key) -and !([FT_Array]::Contains_Empty($_.psobject.Properties.values)))
    }
    return $filled_arr
  }

  static[Bool] Contains_Empty([String]$_arr) {
    $private:result = foreach ($_ in $_arr) { [String]::IsNullOrEmpty($_) }
    return ($True -in $result)
  }

  static [PSCustomObject[]]Incomplete ([PSCustomObject[]]$_arr) {
    $private:incomplete_list = $_arr | Where-Object { [FT_Array]::Contains_Empty($_.psobject.Properties.value) }
    return $incomplete_list 
  }


}

