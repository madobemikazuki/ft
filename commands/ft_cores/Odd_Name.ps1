class Odd_Name {
<#
書類の特定の項目に記載する文字列が規則通り転記するとと冗長すぎるときや、
なぜか半角カタカナが用いられたりする意味不明なデータを相手にするとき、
そのような汚い文字列の代わりとなるより適切な文字列をオブジェクトから参照するメソッドです。
 $_odds として渡すオブジェクトにあらかじめJSONファイルなどで定義していけば、
ソースコードを修正することなく汚い文字列に対処することが期待できます。
#>

  static [String]To_Appropriate([String]$_key, [PSCustomObject]$_odd_dict) {
    $odd_keys = $_odd_dict.psobject.properties.Name
    $result = if ($odd_keys.Contains($_key)) { $_odd_dict.$_key } else { $_key }
    return $result
  }
}

