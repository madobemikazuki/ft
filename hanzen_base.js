"use strict";

const BASE = {
  data() {
    return {
      remarks: [
        {
          title: "**立入許可証番号",
          id: "entry_permit_number",
          source: "",
          digit: 8,
          regexp: /[0-9]{8}/,
          result: "",
          visible: false,
        },
        {
          title: "**管理番号",
          id: "management_number",
          source: "",
          digit: 6,
          regexp: /[0-9]{6}/,
          result: "",
          visible: false,
        },
      ],
    };
  },
  methods: {
    copy(_text) {
      Util.to_Clipboard(_text);
      console.log("copied: " + _text);
    },
    copy_result(_remark) {
      this.copy(_remark.result);
      _remark.result = "";
      _remark.visible = false;
    },
    han_to_zen(_remark) {
      if (this.valid(_remark)) {
        _remark.result = Util.to_zenkaku(_remark.source);
        _remark.source = "";
      }
    },
    valid(_remark) {
      _remark.visible = _remark.regexp.test(_remark.source);
      return _remark.visible;
    },
  },
};
