function read_CSV{
  Param(
    [Parameter(Mandatory=$true)]
    [ValidatePattern("\b\.csv\b$")]
    [String]
    $path
  )
  # [PSCustomObject, PSCustomObject, ...]を返す
  return Import-CSV $path -Encoding Default
}
exit 0

