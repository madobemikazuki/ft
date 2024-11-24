class PoeObject {
  [String]$Name
  [String]$Value
  [int16]$Point_X
  [int16]$Point_Y

  PoeObject([String]$_Name, [String]$_Value, [int16]$_Point_X, [int16]$_Point_Y) {
    #共通情報を転記用のフォーマットに変換する
    $this.Name = $_Name
    $this.Value = $_Value
    $this.Point_X = $_Point_X
    $this.Point_Y = $_Point_Y
    #$this.Point_Column = $_Point_X
    #$this.Point_ROW = $_Point_Y
  }
}


