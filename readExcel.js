const xlsx = require('xlsx');
const calc = require('./calc');
const excel_num = require('./excel-col');

function readIn(book) {
  try {
    xlsx.readFile(book)
  } catch (error) {
    console.log(error);
  }

  let workbook = xlsx.readFile(book);
  let sheet0 = workbook.SheetNames[0];
  let worksheet = workbook.Sheets[sheet0];
  let coordinateX = [];
  let coordinateY = [];
  let coordinate = [];
  let render_data = {};
  let columns = 0;
  let reg = /[0-9]/;
  // 最多支持 500 列数据，并获取当前的列数
  for (i = 0; i < 500; i++) {
    if (worksheet[excel_num.head[i]]) {
      columns ++;
    }
  }

  render_data.fileName = book;
  render_data.coordinate = [];
  render_data.series = [];
  render_data.name = [];

  for (i = 1; i < columns; i++) {
    coordinateY.push([]);
    coordinate.push([]);
    render_data.coordinate.push([]);
    render_data.series.push({});
    console.log(excel_num.head[i]);
    render_data.name.push(worksheet[excel_num.head[i]].v);
  }

  for (let z in worksheet) {
    if (z[0] === '!') continue;
    if (parseInt(z.substring(1)) > 6) {
      let column_number = excel_num.letter_to_number[(z.slice(0, reg.exec(z).index))]
      if (column_number === -1) {
        coordinateX.push(parseFloat(worksheet[z].v));
      } else {
        coordinateY[column_number].push(parseFloat(worksheet[z].v));
      }
    }
  }
  console.log(coordinateX);
  console.log(coordinateY);

  for (let j in coordinateY) {
    for (let k in coordinateX) {
      coordinate[j].push([coordinateX[k], coordinateY[j][k]])
    }
    render_data.series[j] =  {
      name: render_data.name[j],
      type: 'line',
      smooth: true,
      symbol: 'none',
      sampling: 'average',
      data: coordinate[j]
    };
  }

  console.log(coordinate[0])

  for (let i in coordinateY) {
    render_data.coordinate[i] = calc(coordinateX, coordinateY[i]);
  }

  return render_data;
}

/**
 * 用于转换大写字母到自然数，映射关系为 {A-Z} -> {1-26}
 * @param {String} s - 用于转换的大写字母
 * @return {Number} 转换后的自然数
 */
function c2n(s) {
  return s.charCodeAt(0) - 64
}


module.exports = readIn;
