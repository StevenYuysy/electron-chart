const xlsx = require('xlsx');

function datenum(v, date1904) {
	if(date1904) v+=1462;
	let epoch = Date.parse(v);
	return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
}

function sheet_from_array_of_arrays(data, opts) {
	let ws = {};
	let range = {s: {c:10000000, r:10000000}, e: {c:0, r:0 }};
	for(let R = 0; R != data.length; ++R) {
		for(let C = 0; C != data[R].length; ++C) {
			if(range.s.r > R) range.s.r = R;
			if(range.s.c > C) range.s.c = C;
			if(range.e.r < R) range.e.r = R;
			if(range.e.c < C) range.e.c = C;
			let cell = {v: data[R][C] };
			if(cell.v == null) continue;
			let cell_ref = xlsx.utils.encode_cell({c:C,r:R});

			if(typeof cell.v === 'number') cell.t = 'n';
			else if(typeof cell.v === 'boolean') cell.t = 'b';
			else if(cell.v instanceof Date) {
				cell.t = 'n'; cell.z = xlsx.SSF._table[14];
				cell.v = datenum(cell.v);
			}
			else cell.t = 's';

			ws[cell_ref] = cell;
		}
	}
	if(range.s.c < 10000000) ws['!ref'] = xlsx.utils.encode_range(range);
	return ws;
}

function Workbook() {
	if(!(this instanceof Workbook)) return new Workbook();
	this.SheetNames = [];
	this.Sheets = {};
}

function output(data) {
  console.log(data);
  /* original data */
  let o_data = [[1,2,3],[true, false, null, "sheetjs"],["foo","bar",new Date("2014-02-19T14:30Z"), "0.3"], ["baz", null, "qux"]]
  let ws_name = "SheetJS";

  let wb =  xlsx.readFile(data.fileName), ws = sheet_from_array_of_arrays(o_data);

  /* add worksheet to workbook */
  wb.SheetNames.push(ws_name);
  wb.Sheets[ws_name] = ws;

  /* write file */
  /* bookType can be 'xlsx' or 'xlsm' or 'xlsb' */
  var wopts = { bookType:'xlsx', bookSST:false, type:'binary' };

  xlsx.write(wb, wopts);
}

module.exports = output;
