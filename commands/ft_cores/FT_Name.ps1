
class FT_Name {


  static [String] One_Liner ([String[]]$_name_list) {
    $private:blanc = '　'
    $private:under_score = '_'
    $onliner = $_name_list -join $under_score
    return  $onliner.Replace($blanc, "")
  }

  # 上記 One_Liner メソッドのオーバーロードメソッド
  static [String] One_Liner([PSCustomObject[]]$_arr, [String]$_key){
    $private:arr = $_arr
    $private:new_arr = foreach($_ in $arr){
      $_.$_key
    }
    $private:result = [FT_Name]::One_Liner($new_arr)
    return $result
  }

  static [String] Binding ([String]$_first_name, [String]$_last_name, [char]$_delimiter) {
    $sb = New-Object System.Text.StringBuilder
    foreach ($_ in @($_first_name, $_delimiter , $_last_name)) {
      $sb.Append($_)
    }
    return $sb.ToString()
  }


  static [String] Shortened_Com_Type_Name([String] $_corporate_name) {
    $short_name = switch ($_corporate_name) {
      # 半角カッコを全角に変換する。
      # -match オプションなら正規表現が使える
      # -replace オプションなら正規表現が使える
      { $_ -match "株式会社|\(株\)" } { return $_ -replace "株式会社|\(株\)", "（株）"; break }
      { $_ -match "有限会社|\(有\)" } { return $_ -replace "有限会社|\(有\)", "（有）"; break }
      default { return $_ }
    }
    return $short_name
  }


  static [String] Youre_Company_Name ([String]$_management_com_name, [String]$_employer_name) {
    $delimiter = '／'
    $name = if ($_management_com_name -eq $_employer_name) {
      return $_management_com_name
    } if (!($_management_com_name -eq $_employer_name)) {
      return [FT_Name]::Binding($_management_com_name, $_employer_name, $delimiter)
    }
    return $name
  }

  static [String] Replace_Head([String]$_name, [String]$_from, [String]$_to){
    $private:name= $_name
    if ($name -match "^$_from") {
      return $name.Replace($_from, $_to)
    }
    else {
      Write-Host "想定外の値である可能性があります。=> "$name
      return $name
      #Throw ($name + " は 頭文字 " + $_from + " ではないため処理できません。")
    }
  }
}

