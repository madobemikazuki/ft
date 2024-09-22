Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

function Is_Today_or_Future {
  Param
  (
    [Parameter(Mandatory)]
    [DateTime]$_date
  )
  # [DateTime].CompareTo()では実装が複雑化してしまうのでやめた。 
  $private:time_span = (New-TimeSpan (Get-Date) $_date).Days
  ($time_span -gt 0) -or ($time_span -eq 0)
}

# 適切な名前とは言えないな。
function fn_Date_Validation {
  #pester の初期コマンド throw [NotImplementedException]'future_date is not implemented.'
  Param(
    [Parameter(Mandatory)]
    [ValidatePattern("^\d{8}")]
    [String]$_date
  )
  $private:date = [DateTime]::ParseExact($_date, 'yyyyMMdd', $null)
  if (Is_Today_or_Future $date ) {
    return $date
  }
  else {
    Throw "$_date : 昨日よりも未来を。"
  }
}

function Slash {
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory)]
    [ValidatePattern("^\d{8}")]
    [String]$_date
  )
    [DateTime]$private:date = fn_Date_Validation $_date
    return $date.ToString("yyyy/MM/dd")
}

function Blanc_Zen {
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory)]
    [ValidatePattern("^\d{8}")]
    [String]$_date
  )
  [DateTime]$private:date = fn_Date_Validation $_date
  return $date.ToString("yyyy年　MM月　dd日")
}

