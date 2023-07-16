'use strict';
/**
 * 開発者コンソールを開き
 * > FUCKING_TOSHIBA_TESTES.run() でユニットテストを実行
 */


//サンプルデータ
const STRING_WITH_TAB = '作業番号\t作業件名\t\t作業開始日\t\t\t\t\t\t作業終了日\nA190AL\t自社管理業務\t\t2022-10月-01\t\t\t\t\t\t2022-10月-01\n19A136\t１F-１〜四号機 既設多核種除去設備点検手入れ工事(2022)\t\t2022-10月-02\t\t\t\t\t\t2022-10月-02\n2014587\t1F-1〜4号機 多核種除去設備ﾍﾞﾝﾄﾌｨﾙﾀ改良および同関連除却\t\t2022-10月-03\t\t\t\t\t\t2022-10月-03\n210081\t自社管理業務\t\t2022-10月-01\t\t\t\t\t\t2022-10月-01\n201458\t１F-１〜四号機 既設多核種除去設備点検手入れ工事(2022)\t\t2022-10月-02\t\t\t\t\t\t2022-10月-02';
class FUCKING_TOSHIBA_TESTS {

  static run() {
    this.test_Util();
  }

  static test_Util() {
    let result_lines = Util.to_lines(STRING_WITH_TAB, '\n');
    let expected_lines = ['作業番号\t作業件名\t\t作業開始日\t\t\t\t\t\t作業終了日', 'A190AL\t自社管理業務\t\t2022-10月-01\t\t\t\t\t\t2022-10月-01', '19A136\t１F-１〜四号機 既設多核種除去設備点検手入れ工事(2022)\t\t2022-10月-02\t\t\t\t\t\t2022-10月-02', '2014587\t1F-1〜4号機 多核種除去設備ﾍﾞﾝﾄﾌｨﾙﾀ改良および同関連除却\t\t2022-10月-03\t\t\t\t\t\t2022-10月-03', '210081\t自社管理業務\t\t2022-10月-01\t\t\t\t\t\t2022-10月-01', '201458\t１F-１〜四号機 既設多核種除去設備点検手入れ工事(2022)\t\t2022-10月-02\t\t\t\t\t\t2022-10月-02'];
    Test.assert_equals(result_lines, expected_lines, '1 : Util.to_lines()');

    let result_cells = Util.to_cells(STRING_WITH_TAB, '\n');
    let expected_cells = expected_lines;
    Test.assert_equals(result_cells, expected_cells, '2 : Util.to_cells()');

    Test.assert_equals(Util.N_array(3, (_) => true), [true, true, true], '3 : Util.N_array()')
    Test.assert_equals(Util.N_array(10, (_) => 1 + 2), [3, 3, 3, 3, 3, 3, 3, 3, 3, 3], '4 : Util.N_array()');
    Test.assert_equals(Util.N_array(3, (str = 'Hello,') => str + ' world.'), ['Hello, world.', 'Hello, world.', 'Hello, world.'], '5 : Util.Narray()');

    Test.assert_equals(Util.remove_empty(['ok', 'ng', '', 'empty']), ['ok', 'ng', 'empty'], '6: remove_empty()');
    Test.assert_equals(Util.remove_empty([1, , 3, 4, 5]), [1, 3, 4, 5], '7 : remove_empty()');

    Test.assert_equals(Util.to_zenkaku('00123456'), '００１２３４５６');
    Test.assert_equals(Util.to_zenkaku('201354'), '２０１３５４')
  }
}
