
/**
 * 用于获取最接近的数值
 * @param arr {Array} - 一系列坐标
 * @param num {Number} - 用于比较的数值
 */
function getCloset(arr, num) {
  let minValue = Math.abs(arr[0] - num);
  let min = arr[0];
   for (let i in arr) {
     value = Math.abs(arr[i] - num)
     if (value < minValue) {
       minValue = value;
       min = arr[i];
     }
   }
   return min;
}

/**
 * 用于坐标计算
 * @param arrX {Array} - 横坐标
 * @param arrY {Array} - 纵坐标
 */
function calc(arrX, arrY) {
  let arr;
  let lowest;
  let highest;
  let middle;
  let startX;
  let startY;
  let endX;
  let endY;
  let halfWidth;
  let leftarea = [];
  let rightarea = [];
  let coordinate_dict = {};
  
  for (let j=0; j < arrX.length; j++) {
    coordinate_dict[arrY[j]] = arrX[j];
  }
  
  arrY.sort((a, b) => {
    return b - a;
  });

  lowest = arrY[arrY.length-1]

  for (let i of arrY) {
    if (coordinate_dict[i] > coordinate_dict[lowest]) {
      highest = i;
      break;
    }
  }

  middle = (highest + lowest) / 2;
  console.log(middle);
  
  for (let i of arrY) {
    if ((i-1) < middle && (i+1) > middle) {
      if (coordinate_dict[i] < coordinate_dict[lowest]) {
        leftarea.push(i);
      } else {
        rightarea.push(i);
      }
    }
  }
  
  console.log(leftarea);
  console.log(rightarea);
  
  startY = getCloset(leftarea, middle);
  startX = coordinate_dict[startY];
  endY = getCloset(rightarea, middle);
  endX = coordinate_dict[endY];
  halfWidth = endX - startX;
  
  console.log('最低点为：' + coordinate_dict[lowest] + ',' + lowest);
  console.log('峰值坐标为' + coordinate_dict[highest] + ',' + highest);
  console.log('起点光谱为' + startX + ',' + startY);
  console.log('结束光谱为'+ endX + ',' + endY);
  console.log('半高宽为' + halfWidth);
  
  return {
    'lowest': {
      'x': coordinate_dict[lowest],
      'y': lowest
    },
    'highest': {
      'x': coordinate_dict[highest],
      'y': highest
    },
    'start': {
      'x': startX,
      'y': startY
    },
    'end': {
      'x': endX,
      'y': endY
    },
    'halfWidth': halfWidth
  }
      
}

module.exports = calc;