'use strict';
class Test {

  static assert_equals(_result, _expected, _test_name) {
    let test = this.some(_result, _expected);
    if (test) {
      console.log('.');
    } else {
      console.log(_test_name);
      console.assert(
        test,
        '\nresult: ' + _result + '\nexpected: ' + _expected
      );
    }    
  }

  static log(_test, _test_name) {

  }
  
  static to_String(_object) {
    return JSON.stringify(_object, 0, null);
  }

  static some(_result, _expected) {
    return this.to_String(_result) === this.to_String(_expected);
  }

}
