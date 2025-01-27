class FT_Path {

  static CopyToDL ([String]$_from_path, [String]$_destination_path) {
    Write-Host "　From: " $_from_path -BackgroundColor DarkBlue
    $private:file_name = Split-Path $_from_path -Leaf 
    $private:dist_path = ($_destination_path + $file_name)  
    Write-Host "　　To: " $dist_path -BackgroundColor DarkBlue
    Copy-Item -Path $_from_path -Destination $dist_path -Force
  }

  static [String] Fixed_Path ([String]$_folder_path) {
    $private:parsed_folder_path = switch ($_folder_path) {
      #{ $_.Contains("＆") } { $_ -replace "＆", "$"; break; }
      # 実行環境ではこのコードでは動かない？
      { $_.Contains("&") -or $_.Contains("＆") } { $_ -replace "&|＆", "$"; break; }
      
      { $_.Contains("Downloads") } {
        $local_home_folder_path = (Get-Variable | Where-Object { $_.Name -eq "HOME" }).Value; 
        ($local_home_folder_path + $_);
        break; 
      }
      default { $_ }
    }
    return $parsed_folder_path
  }
}

