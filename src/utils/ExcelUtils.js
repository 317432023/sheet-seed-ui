/**
 * excel工具函数
 * @created by kangtengjiao 2022-03-04
 * @lastModified by kangtengjiao 2022-03-04
 * @lastModified by kangtengjiao 2022-03-04
 */
const ExcelUtils = {
  /**
   * 将Excel单元格坐标转换为自然数序号或下标<p>
   *   eg.例如将字符串A1解析成第一行第一列;A2解析成第二行第一列
   * @created by kangtengjiao 2022-03-04
   *
   * @param excelPos excel字符串单元格坐标，eg：A1
   * @param returnIndex 是否返回下标，否则返回自然数序号
   */
  convertExcelPos:function(excelPos, returnIndex) {
    if(typeof excelPos === 'string') {
      let regInt = /^\d+$/
      let regWord = /^[A-Z]+$/i
      // (行号)整数字符串的开始下标
      let numIndex = 0

      const len = excelPos.length
      for(let i = 0; i < len; i++) {
        if( !regWord.test(excelPos[i]) ) {
          numIndex = i
          break
        }
      }
      if( numIndex === 0) {
        console.error('无效的convertExcelPos方法参数: ' + excelPos)
        return null
      }

      // 提取行号字符串
      let rowStr = excelPos.substr(numIndex, len - numIndex)
      if( rowStr.length === 0 || !regInt.test(rowStr)) {
        console.error('无效的convertExcelPos方法参数: ' + excelPos)
        return null
      }
      // 行号
      const rowNo = parseInt(rowStr, 10)


      // 提取列号字符串
      let colStr = excelPos.substr(0, numIndex).toUpperCase()

      // 列号
      const colNo = this.convertExcelCol(colStr)

      return returnIndex ?{
        "rowIndex": rowNo - 1,
        "colIndex": colNo - 1
      } : {rowNo, colNo}

    }
    console.error('无效的convertExcelPos方法参数: ' + excelPos)
    return null
  },

  /**
   * 将excel字母字符串（从A开始）纵坐标转为自然数序号（从1开始）
   * @created by kangtengjiao 2022-03-04
   *
   * @param colStr excel字母字符串纵坐标
   * @returns {number} 自然数序号
   */
  convertExcelCol:function(colStr) {
    const colLen = colStr.length
    let colNo = 0
    for(let i = 0; i < colLen; i ++) {
      colNo += ( ((colStr.charCodeAt(i) - 65) + 1) * Math.pow(26,colLen - i - 1) )
    }
    return colNo
  },
  /**
   * 将excel纵坐标自然数序号（从1开始）转为字母字符串纵坐标（从A开始）
   * @created by kangtengjiao 2022-03-05
   *
   * @param colNo excel字母字符串纵坐标
   * @returns {string}
   */
  reconvertExcelCol:function(colNo) {
    if(colNo <= 0) {
      return ''
    }
    let a = new Array()
    do {
      colNo--
      let n = colNo % 26
      a.push( String.fromCharCode(65 + n) )
      colNo = (colNo - n) / 26
    } while(colNo > 0)

    return a.reverse().join('')
  },

  /**
   * 计算 矩形中心点在excel中的位置
   * @created by kangtengjiao 2022-03-04
   *
   * @param left 横坐标
   * @param top 纵坐标
   * @param width 宽
   * @param height 高
   */
  calCenterLocation:function(left, top, width, height) {
    const x = width/2 + left,
      y = height/2 + top
    return {x, y}
  },


  /**
   * 计算下标在Sheet表单的起始坐标
   * @created by kangtengjiao 2022-03-05
   *
   * @param config LuckySheet 配置（其包含了 行高rowlen 与 列宽columnlen 的 配置）
   * @param index 下标
   * @param calRow
   *   true  - index为行下标，返回结果为起始纵坐标;<br>
   *   false - index为列下标，返回结果为起始横坐标
   */
  calStart:function(config, index, calRow) {
    const cellLenJson = calRow? config.rowlen : config.columnlen
    if(!Number.isInteger(index) || index < 0){
      index = 0
    }
    const no = index + 1
    // 边界线汇总 与 len 长度汇总
    let borders = 0, lenDiff = 0
    for(let i = 0; i < no; i++) {
      let cellLen = cellLenJson[String(i)]
      if(cellLen){
        borders += 1
        lenDiff += cellLen
      }
    }
    lenDiff= lenDiff - cellLenJson[index]
    return borders + lenDiff
  },

  /**
   * 计算两个下标在Sheet表单的坐标范围
   * @created by kangtengjiao 2022-03-05
   *
   * @param config LuckySheet 配置（其包含了 行高rowlen 与 列宽columnlen 的 配置）
   * @param beginIndex 起始下标(包含)
   * @param endIndex 结束下标(包含)
   * @param calRow
   *   true  - 两个index为行下标，返回结果为纵坐标范围;<br>
   *   false - 两个index为列下标，返回结果为横坐标范围
   */
  calRange:function(config, beginIndex, endIndex, calRow) {
    const cellLenJson = calRow? config.rowlen : config.columnlen
    if(beginIndex > endIndex) {
      // 对换两个变量的值（ES6 解构语法）
      [endIndex, beginIndex] = [beginIndex, endIndex];
    }
    let begin = this.calStart(config, beginIndex, calRow),
      end = this.calStart(config, endIndex, calRow)
    let cellLen = cellLenJson[String(endIndex)]

    return {begin, end: end + cellLen}
  },

  /**
   * 判定 坐标点在excel中的位置 是否落在指定的范围内
   * @created by kangtengjiao 2022-03-05
   *
   * @param config LuckySheet 配置（其包含了 行高rowlen 与 列宽columnlen 的 配置）
   * @param location 坐标点(eg. @see ExcelUtils.calCenterLocation)
   * @param beginRowIndex 开始行下标
   * @param endRowIndex 结束行下标
   * @param beginColIndex 开始列下标
   * @param endColIndex 结束列下标
   */
  judgeLocationIn:function(config, location, beginRowIndex, endRowIndex, beginColIndex, endColIndex) {
    const verticalRange = this.calRange(config, beginRowIndex, endRowIndex, true)
    //console.log(" 取得单元格垂直方向的坐标范围："+JSON.stringify(verticalRange))
    if( location.y < verticalRange.begin || location.y > verticalRange.end ) {
      return false
    }
    const horizontalRange = this.calRange(config, beginColIndex, endColIndex, false)
    //console.log(" 取得单元格水平方向的坐标范围："+JSON.stringify(horizontalRange))
    if( location.x < horizontalRange.begin || location.x > horizontalRange.end ) {
      return false
    }
    return true
  },

}

export default ExcelUtils