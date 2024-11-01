﻿
class FT_Name {


  static [String] One_Liner ([String[]]$_name_list) {
    $private:blanc = '　'
    $private:under_score = '_'
    $onliner = $_name_list -join $under_score
    return  $onliner.Replace($blanc, "")
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
}

