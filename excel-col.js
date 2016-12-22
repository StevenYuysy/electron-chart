
/**
 * 用于转换大写字母到自然数，映射关系为 {1-26} -> {A-Z}
 * @param {Number} n - 用于转换的自然数
 * @return {String} 转换后的大写字母
 */
function n2c(n) {
  return String.fromCharCode(64 + n)
}

/**
 * 用于将十进制转换为 26 进制，封装 excel 列数
 * @param {Number} n - 用于转换的十进制数
 # @return {String} s - 转换之后的二十六进制字母
 */

function toNumber26(n) {
  let s = '';
  while (n > 0) {
    let m = n % 26;
    if (m == 0) m = 26;
    s = n2c(m) + s;
    n = (n - m)/26;
  }
  return s;
}

/**
 * 导出需要的数组
 */
let excel_num = {};
excel_num.arr = [];
excel_num.head = [];
excel_num.letter_to_number = {};
for (i=1; i <= 500; i++) {
  excel_num.arr.push(toNumber26(i));
  excel_num.head.push(toNumber26(i)+'1');
  excel_num.letter_to_number[toNumber26(i)] = i - 2; // 返回的是第一列对应-1，第二列对应0
}

module.exports = excel_num;
