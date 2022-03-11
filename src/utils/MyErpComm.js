
import excelutils from "./ExcelUtils"
import luckytool from './MyLuckySheetUtils'
/**
 * ERP common
 * @created by kangtengjiao 2022-03-05
 * @lastModified by kangtengjiao 2022-03-05
 * @lastModified by kangtengjiao 2022-03-05
 */
export default {

  /**
   * 转erp单元格jsonProperty为数组
   * @param jsonProperties
   */
  convertJsonToArray: function(jsonProperties){

    const retA = new Array()

    for(let key in jsonProperties) {
      let obj = {}
      obj['propertyName'] = key
      obj['excelPos'] = jsonProperties[key]

      retA.push(obj)
    }

    return retA
  },

  /**
   * 提取码段
   * @param startCell 开始单元格，eg如 A1
   * @return {} 尺码数组
   */
  getSizeRange: function( startCell ) {
    let sizeRange = new Array()
    if(!startCell) {
      return sizeRange
    }

    let pos = excelutils.convertExcelPos(startCell, true)
    if(!pos) {
      return sizeRange
    }
    let {rowIndex, colIndex } = pos // ES6 解构
    this.$getSizeRangeRecursive(sizeRange, rowIndex, colIndex )

    return { rowIndex, colIndex, sizeRange }
  },

  /**
   * 递归提取码段(仅供 MyErpComm 内部方法getSizeRange调用)
   * @param sizeRange
   * @param rowIndex
   * @param colIndex
   */
  $getSizeRangeRecursive:function(sizeRange, rowIndex, colIndex) {
    let size = window.luckysheet.getCellValue(rowIndex, colIndex)
    if( size ) {
      sizeRange.push( size )
      this.$getSizeRangeRecursive(sizeRange, rowIndex, colIndex+1 )
    }
  },

  /**
   * 根据生产指令单号解析出日期款号（确定鞋图）
   * eg. 把 210426-004A-1 或 210426-004A 解析出 210426-004
   * @param manuOrderNo
   */
  getShoeImageNoByManuNo(manuOrderNo) {
    if(!manuOrderNo) {
      return ''
    }

    const sa = manuOrderNo.split('-')
    const saLen = sa.length
    if( saLen < 2 || saLen > 3 ) {
      console.error('生产指令单号格式不正确，其格式要满足 000000-0000X-0')
      return ''
    }

    // sa[1]
    // return sa[0] +

    let regInt = /^\d+$/
    const len = sa[1].length
    let wordIndex = 0
    for(let i = 0; i < len; i++) {
      if( !regInt.test(sa[1][i]) ) {
        wordIndex = i
        break
      }
    }
    if( wordIndex === 0) {
      console.error('无效的生产指令单号: ' + manuOrderNo)
      return ''
    }

    let shoeImageNo = sa[0]+'-'+sa[1].substring(0, wordIndex)

    console.log( "提取鞋图编号(唯一) ：" + shoeImageNo )

    return shoeImageNo
  },

  /**
   * 提取鞋图NO数组和鞋标集合
   * @param manuOrderNoArray 生产指令单号数组
   * @return {shoeImageNoMap, shoeTongueSet}
   */
  getShoeImageNosAndTongues(manuOrderNoArray = [], manuOrderRowNos = [] ) {

    let shoeTongueSet = new Set()
    let shoeImageNoMap = new Map()

    const _this = this
    manuOrderNoArray.forEach(function( manuOrderNo, index, a ){
      const shoeImageNo = _this.getShoeImageNoByManuNo(manuOrderNo)
      const shoeImageRowNoArray = shoeImageNoMap.get(shoeImageNo) || []
      shoeImageRowNoArray.push(manuOrderRowNos[index])

      shoeImageNoMap.set(shoeImageNo, shoeImageRowNoArray)
      shoeTongueSet.add( manuOrderNo )

    })

    for(let [shoeImageNo, shoeImageRowNoArray] of shoeImageNoMap.entries()){
      shoeImageRowNoArray.sort((a,b)=>{if(a>b)return 1; else if(a<b)return -1;else return 0})
    }
    /*
    for(let value of shoeLabelSet.values()){
      console.log("shoeLabelSet = " + value);
    } */

    return { shoeImageNoMap, shoeTongueSet }

  },

  /**
   * 判断一个数组的数字是否递增连续并且无断档
   * @param a
   */
  isStraight(a=[], result) {
    const aLen = a.length
    if(aLen == 1) {
      return true
    }
    for(let i=1; i<aLen; i++) {
      if( a[i] !== (a[i-1] + 1) ) {
        if(result) {
          result['index'] = i
        }
        return false
      }
    }
    return true
  },

  /**
   * 根据图位置和大小取得图片中心点坐标
   * @created by kangtengjiao 2022-03-07
   *
   * @return {"xxx":{x,y}}
   */
  calcImagesCenterLocation(images) {
    const imagesCenterLocation = {}
    for(let k in images) {
      const image = images[k]
      const { left, top, width, height} = image.default
      const location = excelutils.calCenterLocation(left, top, width, height)
      imagesCenterLocation[k] = location
    }
    return imagesCenterLocation
  },

  /**
   * 提取鞋款图
   * @created by kangtengjiao 2022-03-07
   *
   * @return {} 返回格式如：{"210426-004":"image_10cNo_1646461243131","210426-1476":"image_ai10d_1646461243132"}
   */
  extractShoeStyleImage(config, imagesCenterLocation, shoeImageNoMap, shoeImageColStr ) {
    const shoeImageNoImageKey = {}
    for(let [shoeImageNo, shoeImageRowNoArray] of shoeImageNoMap.entries()){
      let beginCell, endCell
      beginCell = shoeImageColStr + shoeImageRowNoArray[0]
      endCell = shoeImageColStr + shoeImageRowNoArray[shoeImageRowNoArray.length - 1]
      //console.log(shoeImageNo+" 的单元格范围" + beginCell + "," + endCell)
      const beginIndexRange = excelutils.convertExcelPos(beginCell, true)
      //console.log(shoeImageNo+" 的单元格开始下标：" + JSON.stringify(beginIndexRange))
      const endIndexRange = excelutils.convertExcelPos(endCell, true)
      //console.log(shoeImageNo+" 的单元格终止下标：" + JSON.stringify(endIndexRange))

      let inRect = false
      for(let k in imagesCenterLocation) {
        const centerLocation = imagesCenterLocation[k]
        //console.log("开始检查图片: "+ k)
        //console.log("图片中心点坐标: "+ JSON.stringify(centerLocation))
        inRect = luckytool.judgeImageIn(config, centerLocation, beginIndexRange, endIndexRange)
        if(inRect) {
          shoeImageNoImageKey[shoeImageNo] = k
          //console.log("***匹配到图片: "+ k )
          break
        }
      }
    }
    return shoeImageNoImageKey
  },
  /**
   * 提取鞋舌标图
   */
  extractShoeTongueImage(config, imagesCenterLocation, imgStartRowNo, shoeTongueLocation  ) {
    // 按行号排序
    shoeTongueLocation.sort(function(a, b) {
      let s1 = a.location.row, s2 = b.location.row
      if (s1 < s2) return -1
      else if (s1 > s2) return 1
      else return 0
    })

    const shoeImageNoImageKey = {}
    
    let curImgStartRowNoMap = new Map()
    
    shoeTongueLocation.forEach((v,i,a)=>{
      let shoeImageNo = v.shoeTongue
      let {row,column,mc} = v.location

      // 算出 鞋舌标图片 开始行号
      let curImgStartRowNo
      if(curImgStartRowNoMap.has(row)){
        curImgStartRowNo = curImgStartRowNoMap.get(row)
      }else{
        if(curImgStartRowNoMap.size === 0) {
          curImgStartRowNo = imgStartRowNo
        } else {
          let rowNoArray = Array.from( curImgStartRowNoMap.keys() )
          // 正序排序
          rowNoArray.sort((a,b)=>{if(a>b)return 1; else if(a<b)return -1;else return 0})

          console.log("curImgStartRowNoMap.keys()="+JSON.stringify(rowNoArray))

          let lastRowNo = rowNoArray[rowNoArray.length - 1]

          curImgStartRowNo = lastRowNo + 1
        }

        curImgStartRowNoMap.set(row, curImgStartRowNo)
      }

      // 鞋舌标图片 开始单元格下标
      const beginIndexRange = {rowIndex:curImgStartRowNo - 1, colIndex:column}

      // 算出终止列下标
      let endColIndex = column
      if(mc){
        let {r,c,rs,cs} = mc
        endColIndex += (cs - 1)
      }

      // 鞋舌标图片 结束单元格下标
      const endIndexRange = {rowIndex:row - 1, colIndex:endColIndex}

      let inRect = false
      for(let k in imagesCenterLocation) {
        const centerLocation = imagesCenterLocation[k]
        inRect = luckytool.judgeImageIn(config, centerLocation, beginIndexRange, endIndexRange)
        if(inRect) {
          shoeImageNoImageKey[shoeImageNo] = k
          //console.log("***匹配到图片: "+ k )
          break
        }
      }

    })

    return shoeImageNoImageKey
  },

  /**
   * 提取鞋图
   * @param images luckysheet images对象
   * @param imageKeyMap
   */
  makeShoeImage(images, imageKeyMap) {
    const shoeImageNoImage = {}
    for(let key in imageKeyMap) {
      let imageKey = imageKeyMap[key]
      const image = images[imageKey]
      shoeImageNoImage[key] = image.src
    }
    return shoeImageNoImage
  },


}