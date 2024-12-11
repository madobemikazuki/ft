class FT_Dict {


  # オブジェクトの配列からオブジェクトを要素としてHashTableを返す
  static [hashtable] Convert([PSCustomObject[]]$_arr, [String]$_primary_key) {
    $hash = @{}
    foreach ($_ in $_arr) {
      $hash[$_.$_primary_key] = $_
    }
    return $hash
  }

  # HashTable から任意の文字列配列のプロパティのみ取り出し新たなHashTable を返すfilter処理
  static [hashtable] Selective([hashtable]$_dict, [String[]]$_selection) {
    $hash = [ordered]@{}
    foreach ($_primary_key in $_dict.keys) {
      $hash[$_primary_key] = $_dict.$_primary_key | Select-Object -Property $_selection
    }
    return $hash
  }

  #与えられた任意の文字列配列に合致する要素を新たなHashTableに格納して返す。
  static [hashtable] Search([hashtable]$_dict, [String[]]$_primary_keys) {
    $hash = @{}
    foreach ($_primary_key in $_primary_keys) {
      if ($_dict.ContainsKey($_primary_key)) {
        $hash[$_primary_key] = $_dict.$_primary_key
      }
      else{
        Write-Host "$_primary_key : この primary_key は存在しません。" -ForegroundColor Yellow
        continue
      }
    }
    return $hash
  }

  # 与えられた $_primary_keys の全てが $_dict に含まれているか判定。
  # JavaScript の every()みたいなもの
  static [Bool]Every([hashtable]$_dict, [String[]]$_primary_keys) {
    # 正規表現パターンで処理する
    $pattern = "^({0})$" -f ($_dict.keys -join '|')
    $result = foreach ($_key in $_primary_keys) {
      $_key -match $pattern
    }
    return !($false -in $result)
  }
}

