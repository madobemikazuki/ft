'use strict';

const BASE = {
  data() {
    return {
      remarks: [
        {
          title: '**立入許可証番号',
          id: 'entry_permit_number',
          source: '',
          digit: 8,
          regexp: /[0-9]{8}/,
          result: '',
          visible: true,
          error_visible: false,
          error_message: '半角数字を 8 個入力しろ。'
        },
        {
          title: '**管理番号',
          id: 'management_number',
          source: '',
          digit: 6,
          regexp: /[0-9]{6}/,
          result: '',
          visible: true,
          error_visible: false,
          error_message: '半角 数字を 6 個入力しろ。'
        }
      ],
    }
  },
  methods: {
    copy(_text) {
      Util.to_Clipboard(_text);
      console.log('copied: ' + _text);
    },
    han_to_zen(_obj) {
      //console.log('keyup: ' + _obj.source);
      if (_obj.regexp.test(_obj.source)) {
        _obj.result = Util.to_zenkaku(_obj.source);
        _obj.error_visible = false;
        _obj.visible = false;
      }
      else {
        _obj.error_visible = true;
      }
    }
  }
}