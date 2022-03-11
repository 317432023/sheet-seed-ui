<template>
  <div class="c-sheet-demo">

    <div id="sheet-mask" v-show="isShowMask" style="position: absolute;z-index: 1000000;left: 0px;top: 0px;bottom: 0px;right: 0px; background: rgba(255, 255, 255, 0.8); text-align: center;font-size: 40px;align-items:center;justify-content: center;">加载模板中...</div>
    <div style="text-align:right">
      <button id="btn-add" class="btn btn-primary" style=" padding:3px 6px; font-size: 12px; margin-right: 10px;" v-on:click="addSheetHandler">提交</button>
      <button id="btn-biz-add" class="btn btn-danger" style=" padding:3px 6px; font-size: 12px; margin-right: 10px;" v-on:click="bizAddHandler">入库</button>
    </div>
    <div id="luckysheet" style="margin:0px;padding-top:20px;position:absolute;width:100%;height:90%;left: 0px;"></div>


  </div>
</template>

<script>
import LuckyExcel from 'luckyexcel'
import axios from 'axios'
import luckytool from '../utils/MyLuckySheetUtils'
import erpsheets from '../api/erpsheets'

export default {
  name: 'SheetView',
  data() {
    return {
      // 定时器
      timer: undefined,
      // 遮罩
      isShowMask:true,
      // excel 初始参数
      opts : {
        container: 'luckysheet', //容器id
        gridKey: 'ManuOrder', // 唯一key
        title: '订单模板', // 设定表格名称
        lang: 'zh', // 设定表格语言,
        allowEdit: true,//作用：是否允许前台编辑
        showinfobar: false,//作用：是否显示顶部信息栏
      },
      // ERP 表单
      erpsheet: undefined,
    }
  },
  methods:{
    // 初始化
    init(){
      this.erpsheet = erpsheets[this.opts.gridKey]
      if(!this.erpsheet) {
        console.error('不支持模板：'+this.opts.gridKey)
        alert('加载失败')
        return
      }
      this.recovery()
    },
    // 恢复文件
    recovery() {
      let _this = this
      const key = _this.opts.gridKey
      let sheetsData = localStorage.getItem(key)
      if(sheetsData) {
        let title = _this.opts.title
        _this.createSheet(JSON.parse(sheetsData), title, '')
      } else {
        _this.loadRemote()
      }
    },
    // 从远程加载模板文件
    loadRemote() {
      let _this = this
      console.log('开始加载远程模板')

      // 自动加载远程模板
      LuckyExcel.transformExcelToLuckyByUrl(_this.erpsheet.template, _this.opts.title, function(exportJson, luckysheetfile){

        // 文件格式验证
        if(exportJson.sheets==null || exportJson.sheets.length==0){
          alert("读取excel文件内容失败, 当前不支持 .xls 格式的文件!")
          return
        }
        console.log(exportJson, luckysheetfile)
        _this.isShowMask = true
        luckysheet.destroy()

        const sheetsData = exportJson.sheets, 
          title = exportJson.info.name?exportJson.info.name:exportJson.sheets[0].name, 
          userInfo = exportJson.info.creator;

        _this.createSheet(sheetsData, title, userInfo)
      })
    },
    // 创建sheet
    createSheet(sheetsData,title,userInfo) {
      let _this = this
      let payload = {
        data:sheetsData,
        title,
        userInfo, // '<i style="font-size:16px;color:#ff6a00;" class="fa fa-taxi" aria-hidden="true"></i> Lucky'//
        hook:{
          workbookCreateAfter:function(){
            _this.isShowMask = false

            //const dataObject = sheetsData[0]
            //_this.logtestinfo(dataObject)

            _this.timer = setInterval(() =>{
              // 某些定时器操作：保存到本地浏览器
              _this.tempSaveSheetHandler()
            }, 2000);
            // 通过$once来监听定时器，在beforeDestroy钩子可以被清除。
            _this.$once('hook:beforeDestroy', () => {
              clearInterval(_this.timer);
            })
          }
        }
      }
      luckysheet.create( Object.assign({}, _this.opts, payload) )
    },
    // 暂存文件
    tempSaveSheetHandler() {
      const sheetsData = JSON.stringify(luckysheet.getAllSheets())
      //console.log(sheetsData)
      const key = this.opts.gridKey
      localStorage.setItem(key, sheetsData)
      //console.debug('暂存文件成功')
    },
    // 添加提交所有(前端不加工)
    addSheetHandler() {
      let _this = this
      const sheetsData = JSON.stringify(luckysheet.getAllSheets())
      axios.post("/server/add_sheet",
        {
          gridKey: _this.opts.gridKey,
          sheetsData
        }
      ).then(function(response){
        //console.log(response)//成功
        if(response.status == 200) {
          console.log(JSON.stringify(response.data))

        }
      }).catch(function(error){
        console.error(error)//失败
      })
    },
    // 控制台打印测试信息
    logtestinfo(dataObject) {

      //console.log( "单元格A1即 第1列 第1行 的数据 = " + luckysheet.getCellValue(0,0) )
      //console.log( "单元格C7即 第3列 第7行 的数据 = " + luckysheet.getCellValue(6,2) )
      //console.log( "单元格X4即 第24列 第4行 的数据 = " + luckysheet.getCellValue(3,23) )

      console.log( "单元格A1即 第1列 第1行 的数据 = " + luckytool.getCellValue('A1') )
      console.log( "单元格C7即 第3列 第7行 的数据 = " + luckytool.getCellValue('C7') )
      console.log( "单元格X4即 第24列 第4行 的数据 = " + luckytool.getCellValue('X4') )

      const sheetConfig = dataObject.config
      console.log( "excel 列宽配置 = " + JSON.stringify( sheetConfig.columnlen) ) // 第一列下标从0开始，没有值列宽表示0
      console.log( "excel 行高配置 = " + JSON.stringify( sheetConfig.rowlen) ) // 第一行下标从0开始，没有值行高表示0
      console.log( "excel 合并配置 = " + JSON.stringify( sheetConfig.merge) )
      const images = dataObject.images
      Object.keys(images).forEach(function(k){
        console.log( "图片 key = " + k)
        console.log( "图片规格 " + JSON.stringify(images[k].default) )
        console.log( "图片 base64 = " + images[k].src )
      })
    },
    // 添加提交有效的业务数据(经过前端加工提取)
    bizAddHandler() {
      if(!this.erpsheet) {
        alert('暂不支持的表单提交操作')
        return
      }
      const sheetsData = luckysheet.getAllSheets()
      // 文件格式验证
      if(sheetsData==null || sheetsData.length==0){
        alert("Sheet 表单不存在!")
        return
      }
      this.erpsheet.bizAdd(sheetsData[0])
    },
  },
  mounted() {
    this.init()

  }
}
</script>
