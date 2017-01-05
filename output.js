const fs = require('fs');
const Excel = require('exceljs');

function output(dir, data) {
  console.log(data);
  console.log(dir);
  let filePath = dir + "/output-data.xlsx";
  var workbook = new Excel.Workbook();
  var worksheet = workbook.addWorksheet('My Sheet');
  worksheet.columns = [
      { header: '对应文件', key: 'fileName', width: 50},
      { header: '对应曲线', key: 'lineName', width: 20},
      { header: '半高宽', key: 'halfWidth', width: 20},
      { header: '最低点X', key: 'lowestX', width: 25},
      { header: '最低点Y', key: 'lowestY', width: 25},
      { header: '峰值坐标', key: 'highest', width: 25},
      { header: '起点光谱', key: 'start', width: 25},
      { header: '结束光谱', key: 'end', width: 25}
  ];
  for (let i = 0; i < data.length; i++) {
    for (let j = 0; j < data[i].coordinate.length; j++) {
      worksheet.addRow({
        fileName: data[i].fileName,
        lineName: data[i].series[j].name,
        halfWidth: data[i].coordinate[j].halfWidth,
        lowestX: data[i].coordinate[j].lowest.x,
        lowestY: data[i].coordinate[j].lowest.y,
        highest: data[i].coordinate[j].highest,
        start: data[i].coordinate[j].start,
        end: data[i].coordinate[j].end,
      })
    }
  }
  fs.writeFile(filePath, " ", function(err) {
    if(err) {
        return console.log(err);
    }
    console.log("The file was saved!");
});

  workbook.xlsx.writeFile(filePath).then(function() {
      // done
      console.log('file is written');
  });

}

module.exports = output;
