class Odd_Name {
<#
  書類の特定の項目に記載する文字列が規則通り転記するとと冗長すぎるときや、
  なぜか半角カタカナが用いられたりする意味不明なデータを相手にするとき、
  そのような汚い文字列の代わりとなるより適切な文字列をオブジェクトから参照するメソッドです。
   $_odds として渡すオブジェクトにあらかじめJSONファイルなどで定義していけば、
  ソースコードを修正することなく汚い文字列に対処することが期待できます。
#>

  static [String]To_Appropriate([String]$_name, [PSCustomObject]$_odd_names) {
    $private:names = $_odd_names.psobject.properties.Name
    $private:result = if ($names.Contains($_name)) { 
      Write-Host "Odd: $_name"
      $_odd_names.$_name
    } 
    else { 
      Write-Host "No Odd: $_name"
      $_name
    }
    return $result
  }
}
