const electron = require('electron');
const BrowserWindow = electron.BrowserWindow;
const dialog = electron.remote.dialog;

const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

const readIn = require('./readExcel');
const output = require('./output');
const echarts = require('echarts');

let btn = document.getElementById('script');
let outputbtn = document.getElementById('excel');
let alert = document.getElementById('alert');

let tableData;

// Promise 封装异步对象
function showInputDialog() {
  return new Promise((resolve, reject) => {
    dialog.showOpenDialog({
      title:"Select a folder",
      properties: ["openDirectory"]
    }, function (folderPaths) {
      // folderPaths is an array that contains all the selected paths
      if(folderPaths === undefined){
        console.log("No destination folder selected");
        return;
      }else{
        console.log(folderPaths);
        resolve(folderPaths[0]);
      }
    })
  })
}

function showOutputDialog() {
  return new Promise((resolve, reject) => {
    dialog.showOpenDialog({
      properties: ["openDirectory", "createDirectory"]
    }, function (folderPaths) {

      // folderPaths is an array that contains all the selected paths
      if(folderPaths === undefined){
        console.log("No destination folder selected");
        return;
      }else{
        console.log(folderPaths);
        resolve(folderPaths[0]);
      }
    })
  });
}

btn.ondragover = btn.ondrop = (ev) => {
  ev.preventDefault();
}

btn.ondrop = (ev) => {

  let books = [];

  // console.log(ev.dataTransfer.files[0].path);
  ev.preventDefault();
  let filePath = ev.dataTransfer.files[0].path;
  if (checkXLSX(filePath)) {
    books.push(filePath);
  } else {
    travel(filePath, (fileName) => {
      console.log(filePath);
      if (checkXLSX(fileName)) {
        books.push(fileName);
      }
    })
  }
  console.log(books);
  render(books);
}

// 按钮点击运行事件
btn.addEventListener('click', () => {
  let books = [];

  showInputDialog().then((dir) => {
    // 遍历当前目录，找出后缀为 xlsx 的文件夹读取，忽略以 ~ 开头的文件
    travel(dir, (fileName) => {
      // console.log(fileName);

      // windows 系统文件路径
      if (fileName.split('\\') > 0) {
        if (fileName.split('.')[fileName.split('.').length-1] == 'xlsx'
        && fileName.split('\\')[fileName.split('\\').length-1][0] != '~') {
          books.push(fileName);
        }
      } else {
        if (fileName.split('.')[fileName.split('.').length-1] == 'xlsx'
        && fileName.split('/')[fileName.split('/').length-1][0] != '~') {
          books.push(fileName);
        }
      }
    })

    // console.log(books);
    render(books);
  }, false);
}, false);

function render(books) {
  let render_area = document.getElementById('render-area');
  let table_data = [];

  alert.innerHTML = '正在渲染数据';
  render_area.innerHTML = null;

  // 对每个表格进行操作
  let book_index = 0;  // 用于定位不同的表格
  for (let book of books) {
    console.log(book);
    render_data = readIn(book);
    console.log(render_data);
    table_data.push(render_data);

    // 渲染数据
    let div_container = document.createElement('div');
    let div_chart = document.createElement('div');
    let div_chart_sc = document.createElement('div');
    div_container.className = 'chart-container'
    div_chart.className = 'chart';
    div_chart_sc.className = 'chart-sc';
    div_container.appendChild(div_chart);
    div_container.appendChild(div_chart_sc);

    // 表格部分
    let div_table = document.createElement('div');
    div_table.className = 'table';
    div_container.appendChild(div_table);
    let tableHTML = `
    <table>
      <thead>
        <tr>
          <th>对应曲线</th>
          <th>半高宽</th>
          <th>最低点</th>
          <th>峰值坐标</th>
          <th>起点光谱</th>
          <th>结束光谱</th>
        </tr>
      </thead>
    `
    let table_line_index = 0; // 用于区分表格内不同的行数
    for (let i of render_data.coordinate) {
      let className = (table_line_index % 2 == 0) ? 'odd' : 'even';
      tableHTML += `
        <tr class="${className}">
          <td>${render_data.name[table_line_index]}</td>
          <td>${i.halfWidth}</td>
          <td>(${i.lowest.x},${i.lowest.y})</td>
          <td>(${i.highest.x},${i.highest.y})</td>
          <td>(${i.start.x},${i.start.y})</td>
          <td>(${i.end.x},${i.end.y})</td>
        </tr>
      `
      table_line_index ++;
    }
    book_index ++;
    div_table.innerHTML = tableHTML;
    // ecahrt 部分
    render_area.appendChild(div_container);
    let chart = echarts.init(div_chart);
    let chart_sc = echarts.init(div_chart_sc);

    let option = {
      tooltip: {
        trigger: 'axis',
      },
      title: {
        text: render_data.fileName.split('\\')[render_data.fileName.split('\\').length-1],
      },
      legend: {
        top: '10%',
        data: render_data.name
      },
      toolbox: {
        feature: {
          restore: {},
          saveAsImage: {}
        }
      },
      xAxis: {
        name: 'Wavelenggth(nm)',
        nameLocation: 'middle',
        nameGap: '30',
        type: 'value',
        boundaryGap: false
        },
      yAxis: {
        name: 'Transmission(%)',
        nameLocation: 'middle',
        nameGap: '30',
        type: 'value',
        boundaryGap: false
      },
      dataZoom: [{
        type: 'inside',
        start: 50,
        end: 100
      }, {
        start: 0,
        end: 10,
        handleIcon: 'M10.7,11.9v-1.3H9.3v1.3c-4.9,0.3-8.8,4.4-8.8,9.4c0,5,3.9,9.1,8.8,9.4v1.3h1.3v-1.3c4.9-0.3,8.8-4.4,8.8-9.4C19.5,16.3,15.6,12.2,10.7,11.9z M13.3,24.4H6.7V23h6.6V24.4z M13.3,19.6H6.7v-1.4h6.6V19.6z',
        handleSize: '80%',
        handleStyle: {
            color: '#fff',
            shadowBlur: 3,
            shadowColor: 'rgba(0, 0, 0, 0.6)',
            shadowOffsetX: 2,
            shadowOffsetY: 2
        }
      }],
      animation: false,
      series: render_data.series,
    };
    let option2 = {
       legend: {
           right: 10,
       },
       toolbox: {
         feature: {
           restore: {},
           saveAsImage: {}
         }
       },
       tooltip: {
         trigger: 'axis',
         formatter: render_data.result,
       },
       dataZoom: [{
         type: 'inside',
         start: 98,
         end: 100
       }, {
         start: 0,
         end: 10,
         handleIcon: 'M10.7,11.9v-1.3H9.3v1.3c-4.9,0.3-8.8,4.4-8.8,9.4c0,5,3.9,9.1,8.8,9.4v1.3h1.3v-1.3c4.9-0.3,8.8-4.4,8.8-9.4C19.5,16.3,15.6,12.2,10.7,11.9z M13.3,24.4H6.7V23h6.6V24.4z M13.3,19.6H6.7v-1.4h6.6V19.6z',
         handleSize: '80%',
         handleStyle: {
             color: '#fff',
             shadowBlur: 3,
             shadowColor: 'rgba(0, 0, 0, 0.6)',
             shadowOffsetX: 2,
             shadowOffsetY: 2
         }
       }],
       xAxis: {
           splitLine: {
               lineStyle: {
                   type: 'dashed'
               }
           }
       },
       yAxis: {
           splitLine: {
               lineStyle: {
                   type: 'dashed'
               }
           },
           scale: true
       },
       series: {
           data: render_data.scatter,
           type: 'scatter',
           symbolSize: '10',
           label: {
               emphasis: {
                   show: true,
                   position: 'top'
               }
           },
           markLine: {
             animation: false,
             label: {
                 normal: {
                     formatter: render_data.regression.label,
                     textStyle: {
                         align: 'right'
                     }
                 }
             },
             lineStyle: {
                 normal: {
                     type: 'solid'
                 }
             },
             tooltip: {
                 formatter: render_data.regression.label
             },
             data: [render_data.regression.data]
           }
       }
    }
    chart.setOption(option);
    chart_sc.setOption(option2);
  }
  alert.innerHTML = '数据已渲染完毕'
  outputbtn.className = '';
  tableData = table_data;
}

outputbtn.addEventListener('click', event => {
  if (alert.innerHTML === '数据已渲染完毕') {
    console.log(tableData)
  }
}, false)

/**
 * 用于遍历目录下面的文件
 * @param dir {String} - 读取文件名
 * @param callbakc {Function} - 回调函数
 */
 function travel(dir, callback) {
  fs.readdirSync(dir).forEach(function(file) {
    var pathname = path.join(dir, file);

    if (fs.statSync(pathname).isDirectory()) {
      travel(pathname, callback);
    } else {
      callback(pathname);
    }
  })
}

function checkXLSX(fileName) {
  let value = false;
  // windows 系统文件路径
  if (fileName.split('\\') > 0) {
    if (fileName.split('.')[fileName.split('.').length-1] == 'xlsx'
    && fileName.split('\\')[fileName.split('\\').length-1][0] != '~') {
      value = true;
    }
  } else {
    if (fileName.split('.')[fileName.split('.').length-1] == 'xlsx'
    && fileName.split('/')[fileName.split('/').length-1][0] != '~') {
      value = true;
    }
  }
  return value;
}
