Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

$targets = . .\ft_core\io\read_json.ps1 ".\config\FnD.json"
$private:Folder = "${HOME}\Downloads\"
$private:head = "*"
$private:end = "*.*"
foreach ($_ in $targets) {
  $name = ($head + $_ + $end)
  Remove-Item -Path ($folder + $name)
}

