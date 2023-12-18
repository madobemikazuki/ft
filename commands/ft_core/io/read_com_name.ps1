Param(
  [Parameter(Mandatory = $True)]
  [String]$_company_number
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

try {
  $private:source_path = "${HOME}\Downloads\TEMP\"
  $private:json = . .\ft_core\io\read_json.ps1 $source_path
  return $private:json.$_company_number
}
catch [exception] {
  Write-Output "😢😢😢エラーをよく読んでね。"
  $error[0].ToString()
  Write-Output $_
}

exit 0