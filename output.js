const fs = require('fs');
const Excel = require('exceljs');

function output(dir, data) {
  console.log(data);
  console.log(dir);
  let filePath = dir + "/test.xlsx";
  var workbook = new Excel.Workbook();
  var worksheet = workbook.addWorksheet('My Sheet');
  worksheet.columns = [
      { header: '对应文件', key: 'fileName'},
      { header: '对应曲线', key: 'lineName'},
      { header: '半高宽', key: 'halfWidth'},
      { header: '最低点', key: 'lowest'},
      { header: '峰值坐标', key: 'highest'},
      { header: '起点光谱', key: 'start'},
      { header: '结束光谱', key: 'end'}
  ];
  for (let i = 0; i < data.length; i++) {
    for (let j = 0; j < data[i].coordinate.length; j++) {
      worksheet.addRow({
        fileName: data[i].fileName,
        lineName: data[i].series[j].name,
        halfWidth: data[i].coordinate[j].halfWidth,
        lowest: data[i].coordinate[j].lowest,
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
