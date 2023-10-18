<#
.SYNOPSIS
  ftで用いるフォルダと空ファイルを生成する。
.DESCRIPTION
  ローカルにスクリプト群をコピペするのを少しでも楽にできたらなぁと思った。
.EXAMPLE
 PS> . .\cdcf.ps1
.EXAMPLE
  このコマンドレットの使用方法の別の例
.INPUTS
  このコマンドレットへの入力（存在する場合）
.OUTPUTS
  ft 以下のフォルダと.ps1ファイルを出力する。
.NOTES
  生成後の .ps1 ファイルは全て空になっている。
.COMPONENT
  このコマンドレットが属するコンポーネント
.ROLE
  このコマンドレットが属する役割
#>

Set-StrictMode -Version 3.0
#Remove-Item -Path .\ft

# .\tf
New-Item -Path . -Name "ft" -ItemType Directory


# .\ft\commands
New-Item -Path .\ft -Name "commands" -ItemType Directory
$commands_files = @(
  "wbc",
  "ed",
  "gzen"
)
create_dir_with_files_withBOM .\ft\commands $commands_files


# .\ft\commands\ft-core
New-Item -Path .\ft\commands -Name "ft_core" -ItemType Directory
$ft_core_files = @(
  "combined_name",
  "future_date",
  "hanzen",
  "search_regist",
  "your_company_names"
)
create_dir_with_files_withBOM .\ft\commands\ft_core $ft_core_files

# .\ft\commands\ft_core\io\csv
New-Item -Path .\ft\commands\ft_core -Name "io" -ItemType Directory
New-Item -Path .\ft\commands\ft_core\io -Name "csv" -ItemType Directory
create_dir_with_files_withBOM .\ft\commands\ft_core\io\csv @("csv.ps1")

# .\ft\commands\tools
New-Item -Path .\ft\commands -Name "tools" -ItemType Directory
$tools_files = @("pc_info")
create_dir_with_files_withBOM .\ft\commands\tools $tools_files

# .\ft\commands\utils
New-Item -Path .\ft\commands -Name "utils" -ItemType Directory
$utils_files = @("notify", "util_csv", "util_format", "util_txt")
create_dir_with_files_withBOM .\ft\commands\utils $utils_files




function  create_dir_with_files_withBOM {
  Param(
    [Parameter(Mandatory = $true)]
    [String]$folder_path,
    [Parameter(Mandatory = $true)]
    [String[]]$names
  )
  $empty = ""
  $extension = ".ps1"
  foreach ($name in $names) {
    # utf8かつBOM付きのPowerShellスクリプトファイルを生成する。
    $empty | Out-File -FilePath $($folder_path + '\'+ $name +$extension) -Encoding utf8
  }
}