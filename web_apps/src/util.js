class Util {
  static empty = '';

  //空の要素を排除した配列を返す
  static to_lines(_string, _target) {
    return this.separated(_string, _target);
    //return this.remove_empty(array);
  }

  /**
   * this.to_lines()と全く同じだが、
   * [ String, String ] =>  [ [cell,cell,cell], [cell,cell,cell]....]
   * という処理を行う。excell の cellの要素に見立てている。
   * cellが空でもそのまま返す。
     */
  static to_cells(_string, _target) {
    return this.separated(_string, _target);
  }

  static separated(_string, _target) {
    return _string.split(_target);
  }

  static remove_empty(_array) {
    return _array.filter(e => e != this.empty);
  }

  static N_array(N, callback) {
    //Array.fill()はミュータブルなので使えない。
    return [...Array(N)].map(callback);
  }

  //2つのArrayを受取りMapオブジェクトを返す。
  static zip_to_Map(_header, _values) {
    if (_header.length != _values.length) { throw new Error('引数[0] _header と 引数[1] _values 2つの配列の長さが異なるので処理を中止しました') };
    let map = new Map();
    _header.forEach((key, index) => map.set(key, _values[index]));
    return map;
  }

  static to_Clipboard(_text) {
    navigator.clipboard.writeText(_text)
      .then(
        () => {/* clipboard successfully set */
          return;
        },
        () => {/* clipboard write failed */
          return;
        }
      );
  }

  //半角数字文字列を全角数字文字列に変換
  // replace関数は新たな文字列を返す。
  //https://developer.mozilla.org/ja/docs/Web/JavaScript/Reference/Global_Objects/String/replace
  static to_zenkaku(_str) {
    //String.replace() は
    return _str.replace(/[0-9]/g, (s) => String.fromCharCode(s.charCodeAt(0) + 0xFEE0))
  }
}

