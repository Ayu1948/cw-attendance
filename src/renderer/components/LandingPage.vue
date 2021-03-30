<template>
  <div>
    <!-- <input ref="excel-upload-input" class="excel-upload-input" type="file" accept=".xlsx, .xls" @change="handleClick">
    <div class="drop" @drop="handleDrop" @dragover="handleDragover" @dragenter="handleDragover">
      拖拽上传文件
      <el-button :loading="loading" style="margin-left:16px;" size="mini" type="primary" @click="handleUpload">
        点击上传
      </el-button>
    </div>
      <el-table :data="tableData" border highlight-current-row style="width: 100%;margin-top:20px;">
      <el-table-column v-for="item of tableHeader" :key="item" :prop="item" :label="item" />
    </el-table> -->
    <el-switch
      v-model="value1"
      active-text="君尚"
      inactive-color="#13ce66"
      inactive-text="602"
    >
    </el-switch>
    <!-- <el-button @click="">下载</el-button> -->
<a download="table导出Excel" id="excelOut" href="#" rel="external nofollow" >table导出Excel</a>
    <table border="1" id="tableTo602">
      <tr v-for="(item, index) in excelData602" :key="index">
        <td
          width="38"
          v-for="(t, i) in item"
          :key="i"
          :class="index >= 4 && index % 2 != 0 ? 'bw' : 'nw'"
          :colspan="index == 0 ? maxTd : 0"
          :rowspan="index == 0 ? 2 : 0"
          :id="'602-' + index + '-' + i"
          
        >
          <fragment v-if="index >= 4 && index % 2 != 0">
            <fragment v-for="(e, i2) in t">{{ e }}</fragment>
          </fragment>
          <fragment v-else>{{ t }}</fragment>
        </td>

        <th class="else-td" v-if="index==4">姓名</th>
        <th class="else-td" v-if="index==4">0-5min/次</th>
        <th class="else-td" v-if="index==4">5-30min/次</th>
        <th class="else-td" v-if="index==4">补打卡</th>
        <td class="else-td" v-if="newto602[index]">{{newto602[index].name}}</td>
        <td class="else-td" v-if="newto602[index]">{{newto602[index].orange}}</td>
        <td class="else-td" v-if="newto602[index]">{{newto602[index].red}}</td>
        <td class="else-td" v-if="newto602[index]">{{newto602[index].yellow}}</td>
      </tr>
    </table>
    <!-- <table>
      <tr>
      </tr>
      <tr v-for="item in newto602">
        <td>{{item.name}}</td>
        <td>{{item.orange}}</td>
        <td>{{item.red}}</td>
        <td>{{item.yellow}}</td>
      </tr>
    </table> -->
    <table border="1" id="tableToJs">
      <tr v-for="(item, index) in excelDataJS" :key="index">
        <td
          v-for="(t, i) in item"
          :key="i"
          :class="index >= 4 && index % 2 != 0 ? 'bw' : 'nw'"
          :colspan="index == 0 ? maxTd : 0"
          :rowspan="index == 0 ? 2 : 0"
          :id="'js-' + index + '-' + i"
          x:str
        >
          <fragment v-if="index >= 4 && index % 2 != 0">
            <fragment v-for="(e, i2) in t"><br v-if="i2 != 0"/>{{ e }}</fragment>
          </fragment>
          <fragment v-else>{{ t }}</fragment>
        </td>
        
        <th class="else-td" v-if="index==4">姓名</th>
        <th class="else-td" v-if="index==4">0-5min/次</th>
        <th class="else-td" v-if="index==4">5-30min/次</th>
        <th class="else-td" v-if="index==4">补打卡</th>
        <td class="else-td" v-if="newtoJS[index]">{{newtoJS[index].name}}</td>
        <td class="else-td" v-if="newtoJS[index]">{{newtoJS[index].orange}}</td>
        <td class="else-td" v-if="newtoJS[index]">{{newtoJS[index].red}}</td>
        <td class="else-td" v-if="newtoJS[index]">{{newtoJS[index].yellow}}</td>
      </tr>
    </table>
  </div>
</template>

<script>
import xlsx from 'node-xlsx'
// import fs from 'fs'
// const ExcelJS = require('exceljs/dist/es5');
export default {
  name: 'ProgExportImport',
  data() {
    return {
      value1: true,
      fullscreenLoading: false, // 加载中
      imFile: '', // 导入文件el对象
      outFile: '', // 导出文件el对象
      errorDialog: false, // 错误信息弹窗
      errorMsg: '', // 错误信息内容
      excelTitle: {}, // excel标题
      excelData: [], // excel处理数据
      excelData602: [], // excel处理数据
      excelDataJS: [], // excel处理数据
      newto602: null,
      newtoJS: null,
      maxTd: 0
    }
  },
  created() {
    this.getInit()
  },
  mounted() {
    this.maxTd = document.getElementById('tableTo602').rows[3].cells.length
    setTimeout(() => {
      this.newto602.forEach((item, index) => {
        item.list.forEach((t, i) => {
          if (t) {
            // console.log(t, i, index)
            if (t == 'yellow' || t == 'grey')
            document.getElementById('602-' + index + '-' + i).style = `background-color: ${t}`
            else
            document.getElementById('602-' + index + '-' + i).style = `color: ${t}`
          }
        })
      })
      this.newtoJS.forEach((item, index) => {
        item.list.forEach((t, i) => {
          if (t) {
            console.log(t, i, index)
            if (t == 'yellow' || t == 'grey')
            document.getElementById('js-' + index + '-' + i).style = `background-color: ${t}`
            else
            document.getElementById('js-' + index + '-' + i).style = `color: ${t}`
          }
        })
      })
      this.tableToExcel('tableTo602', '下载602')
    }, 1000)
  },
  methods: {
    afterInit() {
      const that = this
      this.newto602 = []
      this.excelData602.forEach((item, index) => {
        if (index >= 4 && index % 2 !== 0) {
          let tmp = [],
            red = 0,
            orange = 0,
            grey = 0,
            yellow = 0
          if (this.excelData602[index - 1][10] == '付丹丹') {
            item.forEach((t, i) => {
              tmp[i] = that.getClassFu(t, index, i)
              if (tmp[i] == 'red') red++
              else if (tmp[i] == 'orange') orange++
              else if (tmp[i] == 'grey') grey++
              else if (tmp[i] == 'yellow') {
                if (yellow < 3) yellow++
                else {
                  tmp[i] = 'red'
                  red++
                }
              }
            })
          } else {
            item.forEach((t, i) => {
              tmp[i] = that.getClass(t, index, i)
              if (tmp[i] == 'red') red++
              else if (tmp[i] == 'orange') orange++
              else if (tmp[i] == 'grey') grey++
              else if (tmp[i] == 'yellow') {
                if (yellow < 3) yellow++
                else {
                  tmp[i] = 'red'
                  red++
                }
              }
            })
          }

          that.newto602[index] = {
            name: this.excelData602[index - 1][10],
            list: tmp,
            red,
            orange,
            grey,
            yellow
          }
        }
      })
      console.log(this.newto602)
      this.newtoJS = []
      this.excelDataJS.forEach((item, index) => {
        if (index >= 4 && index % 2 !== 0) {
          let tmp = [],
            red = 0,
            orange = 0,
            grey = 0,
            yellow = 0
          item.forEach((t, i) => {
            tmp[i] = that.getClass(t, index, i)
            if (tmp[i] == 'red') red++
            else if (tmp[i] == 'orange') orange++
            else if (tmp[i] == 'grey') grey++
            else if (tmp[i] == 'yellow') {
              if (yellow < 3) yellow++
              else {
                tmp[i] = 'red'
                red++
              }
            }
          })
          that.newtoJS[index] = {
            name: this.excelDataJS[index - 1][10],
            list: tmp,
            red,
            orange,
            grey,
            yellow
          }
        }
      })
      console.log(this.newtoJS)
    },
    getClass(arr, row, col) {
      let flag = 0 // 0: 正常 '' 1：5分钟内 orange 2：迟到 red 3: 补卡 yellow -1: 无记录
      const l = arr.length
      const maxTd = this.excelData602[3].length
      switch (l) {
        case 0:
          flag = -1
          break
        case 2: // 1 只有一个
          if (col == maxTd-1) this.lastDay(arr, row, col, false)
          else flag = 3
          break
        default:
          flag = this.judge(arr, row, col)
          break
      }
      if (arr == '') flag = 0
      let t = ''
      switch (flag) {
        case 1:
          t = 'orange'
          break
        case 2:
          t = 'red'
          break
        case 3:
          t = 'yellow'
          break
        case -1:
          t = 'grey'
          break
        default:
          break
      }
      return t
    },
    getClassFu(arr, row, col) {
      let flag = 0 // 0: 正常 '' 1：5分钟内 orange 2：迟到 red 3: 补卡 yellow -1: 无记录
      const l = arr.length
      const maxTd = this.excelData602[3].length
      switch (l) {
        case 0:
          flag = -1
          break
        case 2: // 1 只有一个
          if (col == maxTd-1) this.lastDay(arr, row, col, true)
          else flag = 3
          break
        default:
          flag = this.judgeFu(arr)
          break
      }
      let t = ''
      switch (flag) {
        case 1:
          t = 'orange'
          break
        case 2:
          t = 'red'
          break
        case 3:
          t = 'yellow'
          break
        case -1:
          t = 'grey'
          break
        default:
          break
      }
      return t
    },
    judge(arr, row, col) {
      // 0: 正常 '' 1：5分钟内 orange 2：迟到 red 3: 补卡 yellow -1: 无记录
      const l = arr.length - 1
      let am = arr[0],
        pm = arr[l]
      let tmp1 = am.split(':')
      if (!pm) pm = arr[l - 1]
      let tmp2 = pm.split(':')
      let amLastTime = 9
      if (Number(tmp1[0]) >= 0 && Number(tmp1[0]) < 3) {
        am = arr[1]
        tmp1 = am.split(':')
        amLastTime = 10
      }
      if (Number(tmp2[0]) >= 0 && Number(tmp2[0]) < 3) {
        return -1
      }
      if (Number(tmp1[0]) > amLastTime && Number(tmp1[0]) < 14) {
        return 2
      } else if (Number(tmp1[0]) == amLastTime) {
        if (Number(tmp1[1]) > 36) return 2
        else if (Number(tmp1[1]) > 30) return 1
      }
      if (Number(tmp1[0]) >= 14) {
        return 3
      }
      if (
        (Number(tmp2[0]) >= 14 && Number(tmp2[0]) < 18) ||
        (Number(tmp2[0]) == 18 && Number(tmp2[1]) < 30)
      ) {
        return 2
      } else if (Number(tmp2[0]) < 14) {
        return 3
      }
      return 0
    },
    judgeFu(arr, row, col) {
      // 0: 正常 '' 1：5分钟内 orange 2：迟到 red 3: 补卡 yellow -1: 无记录
      const l = arr.length - 1
      let am = arr[0],
        pm = arr[l]
      console.log(arr)
      let tmp1 = am.split(':')
      if (!pm) pm = arr[l - 1]
      let tmp2 = pm.split(':')
      if (Number(tmp1[0]) >= 0 && Number(tmp1[0]) < 3) {
        am = arr[1]
        tmp1 = am.split(':')
      }
      if (Number(tmp2[0]) >= 0 && Number(tmp2[0]) < 3) {
        return -1
      }
      if (Number(tmp1[0]) > 10 && Number(tmp1[0]) < 14) {
        return 2
      } else if (Number(tmp1[0]) == 10) {
        if (Number(tmp1[1]) > 36) return 2
        else if (Number(tmp1[1]) > 30) return 1
      }
      if (Number(tmp1[0]) >= 14) {
        return 3
      }
      if (
        (Number(tmp2[0]) >= 15 && Number(tmp2[0]) < 18) ||
        (Number(tmp2[0]) == 18 && Number(tmp2[1]) < 30)
      ) {
        return 2
      } else if (Number(tmp2[0]) < 15) {
        return 3
      }
      return 0
    },
    lastDay(arr, row, col, flag) {
      let am = arr[0]
      let tmp1 = am.split(':')
      if (flag) {
        console.log(arr, row, col, flag)
        //fudandan
        let amLastTime = 10
        if (Number(tmp1[0]) >= 0 && Number(tmp1[0]) < 3) {
          am = arr[1]
          tmp1 = am.split(':')
          amLastTime = 11
        }
        if (Number(tmp1[0]) > amLastTime && Number(tmp1[0]) < 14) {
          return 2
        } else if (Number(tmp1[0]) == amLastTime) {
          if (Number(tmp1[1]) > 36) return 2
        }
        return 0
      } else {
        let amLastTime = 9
        if (Number(tmp1[0]) >= 0 && Number(tmp1[0]) < 3) {
          am = arr[1]
          tmp1 = am.split(':')
          amLastTime = 10
        }
        if (Number(tmp1[0]) > amLastTime && Number(tmp1[0]) < 14) {
          return 2
        } else if (Number(tmp1[0]) == amLastTime) {
          if (Number(tmp1[1]) > 36) return 2
          else if (Number(tmp1[1]) > 30) return 1
        }
        return 0
      }
    },
    sliceArray(arr) {
      let newArr = []
      arr.forEach((s, i) => {
        if (s) {
          const reg = /.{5}/g
          let rs = s.match(reg)
          rs.push(s.substring(rs.join('').length))
          newArr[i] = rs
        } else {
          newArr[i] = []
        }
      })
      return newArr
    },
    getInit() {
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
          this.excelData602 = e.data
        }
      })
      this.excelData602.forEach((e, i) => {
        if (i > 4 && i % 2 != 0) {
          this.excelData602[i] = this.sliceArray(e)
        }
      })
      console.log(this.excelData602)
      for (let i = this.excelData602.length - 1; i >= 0; i--) {
        let e = this.excelData602[i]
        const max = this.excelData602[3].length
        if (e.length < max && i !== 0 && i !== 1) {
          e[max - 1] = ''
        }

        if (e.indexOf('君尚') !== -1) {
          // console.log()

          this.excelDataJS.unshift(e, this.excelData602[i + 1])
          this.excelData602.splice(i, 2)
        }
        if (i >= 0 && i <= 3) {
          this.excelDataJS.unshift(e)
        }
      }
      this.afterInit()
      // var table = document.getElementById('tableToExcel')
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
    },
    //base64转码
    base64(s) {
      return window.btoa(unescape(encodeURIComponent(s)));
    },
    //替换table数据和worksheet名字
    format(s, c) {
      return s.replace(/{(\w+)}/g,
        function (m, p) {
          return c[p];
        });
    },
    tableToExcel(tableid, sheetName) {
      var uri = 'data:application/vnd.ms-excel;base64,';
      var template = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"'+
		    'xmlns="http://www.w3.org/TR/REC-html40"><head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet>'
		    +'<x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets>'
		    +'</x:ExcelWorkbook></xml><![endif]-->' +
        
        '</head><body ><table class="tableTo602" x:str>' +
        ' <style type="text/css">' +
        // 'table {' +
        //   'border-collapse: collapse;' +
        //   'table-layout: fixed;' +
        // '}' +
        'tr {' +
          'border-top: 1px solid #000 !important;' +
          'border-bottom: 1px solid #000 !important;' +
        '}' +
        'tr:nth-child(1) {' +
          'font-weight: bold;' +
          'font-size: 30px;' +
        '}' +
        'tr:nth-child(1) td {' +
          'width: 100%;' +
          'text-align: center;' +
        '}' +
        'tr:nth-child(4) td {' +
          'border-right: 1px solid #000;' +
          '/* width: 100%; */' +
          '/* display: block; */' +
          'text-align: center;' +
        '}' +
        '.else-td {' +
          'text-align: center;' +
          'border: none;' +
          'padding: 0 5px;' +
        '}' +
        'td {' +
          'border: none;' +
          'width: 40px !important;' +
          'height: 100%;' +
          '/* display: block; */' +
          'min-height: 25px;' +
          'padding: 0 10px;' +
        '}' +
        'td.orange {' +
          'color: orange;' +
        '}' +
        'td.red {' +
          'color: red;' +
        '}' +
        '.yellow {' +
          'background: #ffff00;' +
        '}' +
        'td.grey {' +
          'background: grey;' +
        '}' +
        'td.bw {' +
          'border-right: 1px solid #000;' +
          '/* div { */' +
          'word-wrap: break-word;' +
          '/* } */' +
        '}' +
        'td.nw {' +
          '/* div { */' +
          'padding: 5px 0;' +
          'white-space: nowrap;' +
          '/* } */' +
        '}' +
        '</style>' +
        '{table}</table></body></html>';
      if (!tableid.nodeType) tableid = document.getElementById(tableid);
      var ctx = {worksheet: sheetName || 'Worksheet', table: tableid.innerHTML};
      document.getElementById("excelOut").href = uri + this.base64(this.format(template, ctx));
      // console.log(uri + this.base64(this.format(template, ctx)))
    }
  }
}
</script>

<!-- Add "scoped" attribute to limit CSS to this component only -->
<style lang="scss" scoped>
table {
  border-collapse: collapse;
  table-layout: fixed;
  tr {
    border-top: 1px solid #000;
    border-bottom: 1px solid #000;
    &:nth-child(1) {
      font-weight: bold;
      font-size: 30px;
      td {
        width: 100%;
        text-align: center;
      }
    }
    &:nth-child(4) {
      td {
        border-right: 1px solid #000;
          /* width: 100%; */
          /* display: block; */
          text-align: center;
      }
    }
  }
  .else-td {
    text-align: center;
    border: none;
    padding: 0 5px;
  }
  td {
    border: none;
    width: 38px !important;
    height: 100%;
    /* display: block; */
    min-height: 25px;
    padding: 0 10px;
    &.orange {
      color: orange;
    }
    &.red {
      color: red;
    }
    &.yellow {
      background: yellow;
    }
    &.grey {
      background: grey;
    }
    &.bw {
      border-right: 1px solid #000;
      /* div { */
        word-wrap: break-word;
      /* } */
    }
    &.nw {
      /* div { */
        padding: 5px 0;
        white-space: nowrap;
      /* } */
    }
  }
}
</style>