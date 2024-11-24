class FT_Date {

  # ft用のDateTime変換コマンド
  # 基本的に今日の日付か未来の日付を返す
  # ftで過去の日付を扱うことはない。
  # Excel に 返す時は文字列に変換する。

  static [Boolean] Validate([String]$_date) {
    $private:result = if ($_date -match "^\d{8}") { $True }
    else { Throw "Error >>> 半角数字を 8文字 入力してください" }
    return $result
  }

  static [Boolean] Is_Today_or_Future([DateTime]$_date) {
    $private:time_span = (New-TimeSpan (Get-Date) $_date).Days
    $private:result = if (($time_span -gt 0) -or ($time_span -eq 0)) { $True }
    else { Throw  "Error >>> 過去のお前には用はない。今日か、今日以降の日付けを入力しろ。" }
    return $result
  }

  static [DateTime] Convert ([string]$_date) {
    $some_day = if ([FT_DATE]::Validate($_date)) { [DateTime]::ParseExact($_date, 'yyyyMMdd', $null) }
    $valid_date = if ([FT_DATE]::Is_Today_or_Future($some_day)) { $some_day }
    return $valid_date
  }

  static [String] Ja_Format ([String]$_date) { 
    $some_date = [FT_Date]::Convert($_date)
    return $some_date.ToString("yyyy年MM月dd日")
  }

  static [String] Ja_Excel_Hell_Format([String]$_date) { 
    $some_date = [FT_Date]::Convert($_date)
    return $some_date.ToString("yyyy年 M月 d日")
  }

  static [String] Ja_Excel_Hell_Full_Format([String]$_date) { 
    $some_date = [FT_Date]::Convert($_date)
    return $some_date.ToString("yyyy年　MM月　dd日")
  }

  static [String] Ja_Empty_Format() {
    return "年　　月　　日"
  }

  static [String] Slash_Format([String]$_date) { 
    $some_date = [FT_Date]::Convert($_date)
    return $some_date.ToString("yyyy/MM/dd")
  }

  static [String] From_Today_Onwards([DateTime]$_today, [DateTime]$_other_day, [String]$_format) {
    $days_span = (New-TimeSpan $_today $_other_day).Days  
    $result = switch ($days_span) {
      {$_ -gt 0} { return $_other_day.ToString($_format); break}
      {$_ -eq 0} { return $_today.ToString($_format); break}
      {$_ -lt 0} { return $_other_day.AddYears(+1).ToString($_format); break}
    }
    return $result
  }
}

