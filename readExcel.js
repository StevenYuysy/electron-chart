const xlsx = require('xlsx');
const calc = require('./calc');
const headArr = require('./excel-col');

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

  // 最多支持 500 列数据，并获取当前的列数
  for (i = 0; i < 500; i++) {
    if (worksheet[headArr[i]]) {
      columns ++;
    }
  }

  render_data.fileName = book;
  render_data.coordinate = [];
  render_data.series = [];
  render_data.name = [];
  render_data.color = [];
  for (i = 1; i < columns; i++) {
    coordinateY.push([]);
    coordinate.push([]);
    render_data.coordinate.push([]);
    render_data.series.push({});
    console.log(headArr[i]);
    render_data.name.push(worksheet[headArr[i]].v);
    render_data.color.push(getRandomColor());

  }
  
  for (let z in worksheet) {
    if (z[0] === '!') continue;
    if (parseInt(z.substring(1)) > 6) {
      if (z[0] == 'A') {
        coordinateX.push(parseFloat(worksheet[z].v));
      } else if (z[0] == 'B') {
        coordinateY[0].push(parseFloat(worksheet[z].v));
      } else if (z[0] == 'C') {
        coordinateY[1].push(parseFloat(worksheet[z].v));
      } else if (z[0] == 'D') {
        coordinateY[2].push(parseFloat(worksheet[z].v));
      } else if (z[0] == 'E') {
        coordinateY[3].push(parseFloat(worksheet[z].v));
      } else if (z[0] == 'F') {
        coordinateY[4].push(parseFloat(worksheet[z].v));
      }
    }
  }

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
      itemStyle: {
          normal: {
              color: render_data.color[j]
          }
      },
      data: coordinate[j]
    };
  }

  console.log(coordinate[0])

  for (let i in coordinateY) {
    render_data.coordinate[i] = calc(coordinateX, coordinateY[i]);
  }

  return render_data;
}

function getRandomColor() {
  var letters = '0123456789ABCDEF';
  var color = '#';
  for (var i = 0; i < 6; i++ ) {
    color += letters[Math.floor(Math.random() * 16)];
  }
  return color;
}

module.exports = readIn;
