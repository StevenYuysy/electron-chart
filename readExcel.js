const xlsx = require('xlsx');
const regression = require('regression');
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
  render_data.scatter = [];
  render_data.regression = {};

  for (i = 1; i < columns; i++) {
    // console.log(excel_num.head[i]);
    coordinateY.push([]);
    coordinate.push([]);
    render_data.coordinate.push([]);
    render_data.series.push({});
    render_data.name.push(worksheet[excel_num.head[i]].v);
    render_data.scatter.push([]);
  }

  for (let z in worksheet) {
    if (z[0] === '!') continue;
    if (z.slice(reg.exec(z).index) > 6) {
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

  // console.log(coordinate[0])

  for (let i in coordinateY) {
    render_data.coordinate[i] = calc(coordinateX, coordinateY[i]);
    render_data.scatter[i] = [parseFloat(render_data.name[i]), render_data.coordinate[i].lowest.x]
  }

  // 线性回归计算
  let result = regression('linear', render_data.scatter);
  render_data.regression.label = result.string;
  render_data.regression.data = [
    {
      coord: [render_data.scatter[0][0],
      render_data.scatter[0][0]*result.equation[0]+result.equation[1]],
      symbol: 'none'
    },
    {
      coord:[render_data.scatter[render_data.scatter.length-1][0],
      render_data.scatter[render_data.scatter.length-1][0]*result.equation[0]+result.equation[1]
      ],
      symbol: 'none'
    }
  ]

  return render_data;
}

module.exports = readIn;
