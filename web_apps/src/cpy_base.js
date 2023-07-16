const NEW_LINE = '\n';
const TAB = '\t';
const ZENB = '　';

const BASE = {
  data() {
    return {
      head_line: [],
      applicant: new Map(),
      applicant_name_list: [],
      selected_name: '',
      selected_name_kana:'',
      applicant_list: [],
      displayable: true,
      error_message: ''
    }
  },
  methods: {
    conversion_with_header(_event) {//項目の行と値の行を分離してる。
      let arrays = this.to_arrays(_event.target.value);
      this.head_line = arrays[0];

      try {
        this.applicant_list = arrays
          .filter((_, index) => 0 < index)

          //セルの中身に改行文字列があれば、処理は停止する。
          .map(line => Util.zip_to_Map(this.head_line, line));
      } catch (e) {
        this.error_message = e.message;
        console.dir(arrays);
        console.error(e);
      }

      this.applicant_name_list = this.applicant_list.map(e => this.full_name(e));
      this.displayable = false;

      //あとで消す
      //console.dir(this.head_line);
    },
    to_arrays(_string) {
      let array = Util.to_lines(_string, NEW_LINE);
      return array.map(e => Util.to_cells(e, TAB));
    },
    copy(_text) {
      Util.to_Clipboard(_text);
      console.log('copied: '+ _text);
    },
    exist_lines() {
      return this.applicant_list.length > 0;
    },
    full_name(_Map) {
      return _Map.get('漢字氏名（姓）') + ZENB + _Map.get('漢字氏名（名）');
    },
    full_name_kana(_Map) {
      return _Map.get('カナ氏名（姓）') + ZENB + _Map.get('カナ氏名（名）');
    },
    show_applicant_info(_index) {
      //this.displayable = !this.displayable;
      this.applicant = this.applicant_list[_index];
      this.selected_name = this.full_name(this.applicant);
      this.selected_name_kana = this.full_name_kana(this.applicant)
      /**
      for ([key, value] of this.applicant_list[_index].entries()) {
        console.log(key + ':' + value);
      }
      */
    }
  },

  computed: {
    get_lines() {
      return this.lines;
    },
    get_head_lines() {
      return this.head_line;
    }
  }
};
