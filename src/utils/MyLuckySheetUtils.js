import excelutils from './ExcelUtils'

/**
 * excel工具函数
 * @created by kangtengjiao 2022-03-04
 * @lastModified by kangtengjiao 2022-03-04
 * @lastModified by kangtengjiao 2022-03-05
 */
const MyLuckSheetUtils = {
  /**
   * 提取LuckySheet单元格内容
   * @created by kangtengjiao 2022-03-04
   *
   * @param excelPos Excel形式的字符串单元格坐标
   * @param setting 参考LuckySheet api文档
   */
  getCellValue : function(excelPos, setting) {
    let pos = excelutils.convertExcelPos(excelPos, true)
    if(!pos) {
      return null
    }
    let {rowIndex, colIndex } = pos // ES6 解构
    return window.luckysheet.getCellValue(rowIndex, colIndex, setting)
  },

  /**
   * 批量提取单元格内容
   * @created by kangtengjiao 2022-03-05
   *
   * @param excelPositions 属性坐标列表，格式如[{propertyName:'', excelPos:'', setting}]
   */
  getCellValues: function(excelPositions) {
    let _this = this
    const jsonObj = {}
    excelPositions.forEach((v,i,a)=>{
      let {propertyName, excelPos, setting} = v // ES6 解构
      if(propertyName && excelPos) {
        jsonObj[propertyName] = _this.getCellValue(excelPos, setting)
      }
    })
    return jsonObj
  },

  /**
   * 判断图片的中心点坐标是否落在单元格区域范围内
   * @created by kangtengjiao 2022-03-05
   *
   * @param config LuckySheet 配置（其包含了 行高rowlen 与 列宽columnlen 的 配置）
   * @param imgMeta 图片元素据：坐标与宽高，eg,格式 {left, top, width, height}
   * @param beginCell 起始单元格，eg,格式 "A1"
   * @param endCell 结束单元格，eg,格式 "A1"
   */
  /*judgeImageIn0(conifg, imgMeta, beginCell, endCell) {
    let { left, top, width, height} = imgMeta

    const location = excelutils.calCenterLocation(left, top, width, height)
    console.log("取得图片中心点坐标："+JSON.stringify(location) )
    const beginIndexRange = excelutils.convertExcelPos(beginCell, true)
    console.log("取得单元格开始下标：" + JSON.stringify(beginIndexRange))
    const endIndexRange = excelutils.convertExcelPos(endCell, true)
    console.log("取得单元格终止下标：" + JSON.stringify(endIndexRange))

    return excelutils.judgeLocationIn( conifg, location,
      beginIndexRange.rowIndex, endIndexRange.rowIndex,
      beginIndexRange.colIndex, endIndexRange.colIndex
    )
  },*/

  /**
   * 判断图片的中心点坐标是否落在单元格区域范围内
   * @created by kangtengjiao 2022-03-05
   *
   * @param config LuckySheet 配置（其包含了 行高rowlen 与 列宽columnlen 的 配置）
   * @param imageCenterLocation 图片的中心点坐标 {x,y}
   * @param beginCell 起始单元格下标，eg,格式 {rowIndex, colIndex}
   * @param endCell 结束单元格下标，eg,格式 {rowIndex, colIndex}
   */
  judgeImageIn(config, imageCenterLocation, beginIndexRange, endIndexRange) {
    return excelutils.judgeLocationIn( config, imageCenterLocation,
      beginIndexRange.rowIndex, endIndexRange.rowIndex,
      beginIndexRange.colIndex, endIndexRange.colIndex
    )
  }

}

export default MyLuckSheetUtils