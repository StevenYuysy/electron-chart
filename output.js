const fs = require('fs');
const Excel = require('exceljs');

function output(dir, data) {
  console.log(data);
  console.log(dir);
  let filePath = dir + "/test.xlsx";
  var workbook = new Excel.Workbook();
  var worksheet = workbook.addWorksheet('My Sheet');
  worksheet.columns = [
      { header: 'Id', key: 'id', width: 10 },
      { header: 'Name', key: 'name', width: 32 },
      { header: 'D.O.B.', key: 'DOB', width: 10 }
  ];
  worksheet.addRow({id: 1, name: 'John Doe', dob: new Date(1970,1,1)});
  worksheet.addRow({id: 2, name: 'Jane Doe', dob: new Date(1965,1,7)});
  // workbook.commit();
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
