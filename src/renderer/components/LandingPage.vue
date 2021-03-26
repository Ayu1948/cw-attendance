<template>
  <!-- <input
    id="upload"
    type="file"
    @change="importfxx(this)"
    accept=".csv, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel"
  /> -->
  <table border="1" id="tableToExcel">
    <tr v-for="(item, index) in excelData" :key="index">
      <td v-for="(t, i) in item" :key="i">{{t}}</td>
    </tr>
  </table>
</template>

<script>
import xlsx from 'node-xlsx'
// import fs from 'fs'
// const ExcelJS = require('exceljs/dist/es5');
export default {
  name: 'ProgExportImport',
  data() {
    return {
      fullscreenLoading: false, // 加载中
      imFile: '', // 导入文件el对象
      outFile: '', // 导出文件el对象
      errorDialog: false, // 错误信息弹窗
      errorMsg: '', // 错误信息内容
      excelTitle: {}, // excel标题
      excelData: [], // excel处理数据
      maxTd: 0
    }
  },
  created() {
    this.getInit()
  },
  methods: {
    async getInit() {
      // const workbook = new ExcelJS.Workbook();
      // workbook.xlsx.readFile(`${__dirname}/../assets/goal.xls`)
      //   .then(function() {
      //     const ws = workbook.getWorksheet(2)
      //     const row = ws.getRow(6)
      //     console.log(row)
      //   })
      const workSheetsFromFile = xlsx.parse(`${__dirname}/../assets/goal.xls`)
      console.log(workSheetsFromFile)
      workSheetsFromFile.forEach((e) => {
        if (e.name == '考勤记录') {
          console.log(e.data)
          this.excelData = e.data
        }
      })
      var table = document.getElementById('tableToExcel')
      // table.rows[0]
      // xlsx.utils.sheet_to_html(workSheetsFromFile)
      // let optionsData = []
      // for (let i = 0; i < 31; i++) {
      //   optionsData.push({wch: 3.6, border: 'none'})
      // }
      // const range = {s: {c: 0, r:0 }, e: {c:30, r:1}};
      // const options = {'!cols': optionsData, '!merges': [ range ]};
      // const buffer = xlsx.build([{name: 'test', data: this.excelData}], options)
      // fs.writeFileSync('test.xlsx', buffer, {flag: 'w'})
    }
  }
}
</script>

<!-- Add "scoped" attribute to limit CSS to this component only -->
<style lang="scss" scoped>
table {
  tr {
    border-bottom: 1px solid #000;
  }
  td {
    border: none;
  }
}
</style>