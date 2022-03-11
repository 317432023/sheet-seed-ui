import luckytool from '../../utils/MyLuckySheetUtils'
import excelutils from '../../utils/ExcelUtils'
import comm from '../../utils/MyErpComm'

const sheet = {
  // 模板
  template: 'http://127.0.0.1:8080/'+encodeURIComponent('订单明细 OrderTemplateFile.xlsx'),
  // 元素据，用于构造 业务添加接口 参数
  meta: {
    // 关键字
    keyword: {
      // 明细列表结束 关键字
      detail_end: '总计',
    },
    // 主记录元素据（固定单元格坐标）
    main: {
      pos:{// 已知坐标
        "customer": "A1", // 客人单元格坐标
        "size_range_start": "K4", // 尺码范围开始单元格坐标
      },
      // 行号固定，但列位置需要通过计算获得
      rows: {
        "order_date" : 2, // 下单日期 位于第2行
      },
      // 列号固定，但行位置需要通过计算获得
      cols: {
        // ...
        "description": 'B', // 备注 位于第B列
        "tongue_manu_order_no": "K", // 鞋舌标 的图片对应的生产指令单号 起始列 位于第 K 列
      },
    },

    // 明细列表 元素据
    detail: {
      // 明细列表开始单元格坐标 行位置可以通过它 循环计算获得
      start: 'B6',
      // 已知的明细列号
      known_header: {
        "manu_order_no": 'B', // 生产指令单 位于第B列
        "order_no": 'C', // 订单号 位于第C列
        "buyer_shape_name": 'D', // 客人型体名 位于第D列
        "buyer":'E', // 买主
        "upper_material":'F', // 帮面材料
        "outsole_plant":'G', // 大底是否植布
        "color":'H', // 颜色
        "shoe_image":'I', // 鞋图
        "print_code":'J', // 印刷代码

      },
      // 未明确明细列号，需要计算尺码范围与 header 获得
      unspecified_header: {
        "pairs": '', // 总数 PAIRS
        "packing_code": '', // 包装代码
        "packing": '', // 包装方式
        "inner_label_color": '', // 内盒贴标颜色
        "sock_material": '', // 面衬材质
        "test_sample":'',// 测试样
        "tongue":'',// 鞋舌标
        "delivery_date":'',// 交期
      },
    },
    // 间隔数
    intervals: {
      // 明细列表结束（总计）行 与 备行注 的间隔(行)数
      "detail_end_between_description":2,
    }
  },
  /**
   * 业务录入sheet添加
   * @param dataObject
   */
  bizAdd:function(dataObject) {
    const {config, images} = dataObject

    const {keyword, main, detail, intervals} = this.meta

    // 一、构造 newMainPos
    const newMainPos = {}

    // 客户
    newMainPos["customer"] = main.pos["customer"]
    // 码段开始单元格
    const sizeRangeStartCell = main.pos["size_range_start"]
    // 提取尺码范围对象
    const sizeRangePos = comm.getSizeRange(sizeRangeStartCell)
    let {rowIndex, colIndex, sizeRange} = sizeRangePos
    // 尺码个数
    const sizeRangeCounter = sizeRange.length
    // 订单日期、总数量PAIRS的列下标
    let calStartColIndex = colIndex + sizeRangeCounter
    let variableColNo = calStartColIndex + 1

    // 二、提取 明细列表
    let {start, known_header, unspecified_header} = detail

    let {rowNo, colNo} = excelutils.convertExcelPos(start)
    for(let key in unspecified_header) {
      unspecified_header[key] = excelutils.reconvertExcelCol(variableColNo ++)
    }
    // 计算明细标题所在的列
    let header = Object.assign({}, known_header, unspecified_header)

    // 提取列表数据
    let currentRowNo = rowNo
    let endKeyWord = keyword.detail_end
    let endWord = window.luckysheet.getCellValue(currentRowNo - 1, colNo - 1)
    const excelPositions = new Array()
    const detailList = new Array() // 列表数据
    while(endWord && endWord !== endKeyWord){

      // begin 提取 detail
      for(let key in header) {
        let excelPos = header[key] + currentRowNo
        excelPositions.push( { propertyName:key, excelPos } )
      }
      let detailObject = luckytool.getCellValues( excelPositions )
      detailList.push(detailObject)
      // end 提取 detail

      // begin 提取 明细尺码列表
      let detailSizeList = new Array() // 列表数据（尺码）
      let currentSizeColIndex = colIndex
      for(let i = 0; i < sizeRangeCounter; i++) {
        let sizeKey = sizeRange[i]
        let sizeVal = window.luckysheet.getCellValue(currentRowNo - 1, currentSizeColIndex++)
        if(sizeVal){
          //detailSizeList.push(sizeVal)
          detailSizeList.push({sizeKey,sizeVal})
        }
      }
      detailObject['detailSizeList'] = detailSizeList
      // end 提取 明细尺码列表

      currentRowNo ++
      endWord = window.luckysheet.getCellValue(currentRowNo - 1, colNo - 1)
    }

    // 三、完善 主数据
    const mainData = this.getMainValues(newMainPos)
    // 设置尺码范围
    mainData['size_range'] = sizeRangePos.sizeRange.join(',')
    // 设置订单日期
    mainData['order_date'] = window.luckysheet.getCellValue(main.rows['order_date'] - 1, calStartColIndex)

    // 提取备注
    const descrColNo = excelutils.convertExcelCol( main.cols['description'] ) // 备注 列
    let currentDescrRowNo = currentRowNo + intervals["detail_end_between_description"] // 备注 起始 行
    const rowsCount = 7 + detailList.length // 迭代备注的行数
    let description = ''
    for(let i=0; i< rowsCount; i++) {
      let descr = window.luckysheet.getCellValue(currentDescrRowNo - 1, descrColNo - 1)
      if(descr) {
        description = description + descr + '\n'
      }
      currentDescrRowNo ++
    }
    mainData['description'] = description

    // 四、提取鞋图
    const manuOrderNos = [], manuOrderRowNos = []
    detailList.forEach(function(v, i, a){
      manuOrderNos.push(v['manu_order_no'])
      manuOrderRowNos.push(rowNo+i)
    })
    let shoeImageNosAndTongues = comm.getShoeImageNosAndTongues(manuOrderNos, manuOrderRowNos)
    let { shoeImageNoMap, shoeTongueSet } = shoeImageNosAndTongues

    if( !this.validateStraight(shoeImageNoMap) ) {
      return
    }

    console.log( "鞋图与鞋标: ")
    for(let [shoeImageNo, shoeImageRowNoArray] of shoeImageNoMap.entries()){
      console.log("鞋图：" + shoeImageNo + ", 对应生产指令单记录行号：" + shoeImageRowNoArray);
    }
    console.log( "鞋标： " + Array.from(shoeTongueSet));
    console.log( "主订单提取结果 : " + JSON.stringify(mainData) )
    console.log( "订单明细提取结果 : " + JSON.stringify(detailList) )

    // 图片中心点坐标
    const imagesCenterLocation = comm.calcImagesCenterLocation(images)

    // 五、提取鞋款图
    const shoeImageColStr = known_header['shoe_image']
    const shoeImageNoImageKey = comm.extractShoeStyleImage(config, imagesCenterLocation, shoeImageNoMap, shoeImageColStr)

    // 数据完整性检查
    for(let [shoeImageNo, shoeImageRowNoArray] of shoeImageNoMap.entries()){
      if( !shoeImageNoImageKey[shoeImageNo] ) {
        alert("未找到"+shoeImageNo+"对应的鞋图，生产指令单号记录行：" + shoeImageRowNoArray)
        return
      }else {
        console.log("找到"+shoeImageNo+"对应的鞋图，生产指令单号记录行：" + shoeImageRowNoArray)
      }
    }

    // 提取图片
    let shoeImageNoImage = comm.makeShoeImage(images, shoeImageNoImageKey)

    console.log("提取鞋款图片结果：")
    for(let key in shoeImageNoImage) {
      let imgDat = shoeImageNoImage[key]
      console.log(key+"="+imgDat)
    }

    // 六、提取鞋舌标图
    // 鞋舌标起始列
    const startTongueColStr = main.cols['tongue_manu_order_no']
    const startTongueColNo = excelutils.convertExcelCol(startTongueColStr)
    //
    /**
     *
     [
       {
        "shoeTongue":"210426-004A",
        "location":{
          "mc": {
            "r": 29,
            "c": 10,
            "rs": 1,
            "cs": 4
          },
          "row": 29,
          "column": 10
        },
        ...
       },
     ]
     */
    const shoeTongueLocation = []
    for(let shoeTongue of shoeTongueSet.values()){
      let location = null
      const manuOrderNoCells = window.luckysheet.find(shoeTongue, {isWholeWord:true})
      manuOrderNoCells.forEach(function (obj) {
        let { mc, row, column} = obj
        if( ( row + 1 ) > currentRowNo && (column + 1) >= startTongueColNo ) {
          location = { mc, row, column}
        }
      })
      shoeTongueLocation.push({shoeTongue, location })
    }

    // 数据完整性检查
    for(let i=0, len = shoeTongueLocation.length; i<len; i++) {
      if(!shoeTongueLocation[i].location) {
        alert("未找到单号"+shoeTongueLocation[i].shoeTongue+"对应的鞋舌标")
        break
      }
    }

    /*shoeTongueLocation.forEach(function (v,i,a) {
      console.log(JSON.stringify(v))
    })*/

    // imgStartRow 初始值为 currentRowNo + 1
    const shoeImageNoImageKey2 = comm.extractShoeTongueImage(config,imagesCenterLocation, currentRowNo + 1, shoeTongueLocation)

    // 数据完整性检查
    for(let shoeTongue of shoeTongueSet.values()){
      if( !shoeImageNoImageKey2[shoeTongue] ) {
        alert("未找到"+shoeTongue+"对应的鞋标图")
        return
      }else {
        console.log("找到"+shoeTongue+"对应的鞋图")
      }
    }

    // 提取图片
    let shoeImageNoImage2 = comm.makeShoeImage(images, shoeImageNoImageKey2)

    console.log("提取鞋舌图片结果：")
    for(let key in shoeImageNoImage2) {
      let imgDat = shoeImageNoImage2[key]
      console.log(key+"="+imgDat)
    }
  },

  /**
   * 提取主表单元格的值 newMainValues（Json格式）
   * @param newMainPos 格式 {'customer':'A1', ...}
   * @return 格式 {'customer':'Levis', ...}
   */
  getMainValues:function(newMainPos) {
    let newMainPosArr = comm.convertJsonToArray( newMainPos )
    let newMainValues = luckytool.getCellValues( newMainPosArr )
    return newMainValues
  },

  validateStraight(shoeImageNoMap) {
    for(let [shoeImageNo, shoeImageRowNoArray] of shoeImageNoMap.entries()){
      const result = {}
      if(!comm.isStraight(shoeImageRowNoArray, result)) {

        const colStr = this.meta.detail.known_header['manu_order_no'], rowNo = shoeImageRowNoArray[ result['index'] ]
        const cellStr = colStr + rowNo
        alert( "行号不连续：" + JSON.stringify(shoeImageRowNoArray) + "，\n原因：相同图片的生产指令单记录要调整到相邻的行记录中" )

        return false
      }
    }
    return true
  },
}

export default sheet