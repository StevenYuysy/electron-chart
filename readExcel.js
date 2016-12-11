const xlsx = require('xlsx');
const calc = require('./calc');

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
  let coordinateY = [[], [], [], [], []];
  let coordinate = [[], [], [] ,[], []];
  let render_data = {};
  render_data.fileName = book;
  render_data.name = [worksheet['B1'].v, worksheet['C1'].v, worksheet['D1'].v, worksheet['E1'].v, worksheet['F1'].v];
  render_data.series = [{}, {}, {}, {}, {}];
  render_data.color = ['#FF7164', '#69FF82', '#94BCFF', '#CCC941', '#FFE364'];
  render_data.coordinate = [[], [], [], [], []];
  
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

module.exports = readIn;