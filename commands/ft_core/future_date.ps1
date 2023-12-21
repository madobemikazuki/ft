<#
.SYNOPSIS
	明日以降の日付を指定したフォーマットで返す
.DESCRIPTION
	このスクリプトは明日以降の日付を指定したフォーマットで返す
.PARAMETER arg
	数字8桁の文字列
.PARAMETER future
	明日以降の日付
.EXAMPLE
	PS> ./future_date.ps1 20221231 | ja_format
.LINK
 未定義
.NOTES
	Author: madobe_mikazuki
#>


# 今はこれでいい
Set-StrictMode -Version 3.0



function future_date {
  #pester の初期コマンド throw [NotImplementedException]'future_date is not implemented.'
  [CmdletBinding()]
  Param
  (
    [Parameter(Mandatory)]
    [ValidatePattern("^\d{8}")][String]$arg
  )
  try
  {
    $date = [DateTime]::ParseExact($arg, 'yyyyMMdd', $null)
    if(is_future $date )
    {
      return $date
    }else{
      Throw
    }
  }
  catch
  {
    Throw "今日より明日だ。覚えておけ。"
  }

}

function is_future{
  Param
  (
    [Parameter(Mandatory)]
    [DateTime]$future
  )
  $today = Get-Date
  $future -gt $today
}

function ja_format([Parameter(ValueFromPipeline=$true)] $future){
  process {
    return $future.ToString("yyyy年MM月dd日");
  }
}

function excel_hell_format([Parameter(ValueFromPipeline=$true)] $future){
  process {
    return $future.ToString("yyyy年　　MM月　　dd日");
  }
}

function slash_format([Parameter(ValueFromPipeline=$true)] $future){
  process {
    return $future.ToString("yyyy/MM/dd");
  }
}

