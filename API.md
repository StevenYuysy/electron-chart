# readExcel 模块

### readIn 函数
- accept: {Array} - 带有目标路径文件，且以 xlsx 结尾的数组
- return: {Object} - 需要渲染的对象，包括
  - fileName {String} - 文件名
  - coordinate {Array} - 继承 calc 模块
  - series {Array} - 用于渲染 Echart 折线的数组
  - name {Array} - 对应曲线名字的数组
  - scatter {Array} - 用于渲染 Echart 散点图的数组
  - result - 线性回归计算式
