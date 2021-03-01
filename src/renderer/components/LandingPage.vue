<template>
  <div id="wrapper">
    <el-upload
　　　　action="/"
　　　　:on-change="uploadChange"
　　　　:show-file-list="false"
　　　　accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
　　　　:auto-upload="false">
　　　　<el-button size="small" icon="el-icon-upload" type="primary">导入数据</el-button>
　　</el-upload>
  </div>
</template>

<script>
import SystemInformation from './LandingPage/SystemInformation'
import FileSaver from 'file-saver'
import XLSX from 'xlsx'
export default {
  name: 'landing-page',
  components: { SystemInformation },
  data() {
    return {
      xlscList: [],
      xlscTitle: {
        "类型": "type",
        "ID": "id",
        "名称": "name"
      },
    }
  },
  methods: {
    uploadChange(file) {
      let self = this;
      const types = file.name.split('.')[1];
      const fileType = ['xlsx', 'xlc', 'xlm', 'xls', 'xlt', 'xlw', 'csv'].some(item => {
        return item === types
      });
      if (!fileType) {
        this.$message.error('文件格式错误，请重新选择文件！')
      }

      this.file2Xce(file).then(tab => {
        // console.log(tab)
　　　　　// 过滤，转化正确的JSON对象格式
        if (tab && tab.length > 0) {
          tab[0].sheet.forEach(item => {
            let obj = {};
            for (let key in item) {
              obj[self.xlscTitle[key]] = item[key];
            }
            self.xlscList.push(obj);
          });
          // console.log(self.xlscList)

          if (self.xlscList.length) {
            this.$message.success('上传成功')
　　　　　　　　// 获取数据后，下一步操作
          } else {
            this.$message.error('空文件或数据缺失，请重新选择文件！')
          }
        }
      })
    },
　　　// 读取文件
　　 file2Xce(file) {
      return new Promise(function(resolve, reject) {
        const reader = new FileReader();
        reader.onload = function(e) {
          const data = e.target.result;
          this.wb = XLSX.read(data, {
            type: "binary"
          });
          const result = [];
          this.wb.SheetNames.forEach(sheetName => {

            result.push({
              sheetName: sheetName,
              sheet: XLSX.utils.sheet_to_json(this.wb.Sheets[sheetName])
            })
          })
          resolve(result);
        }
        reader.readAsBinaryString(file.raw);
      })
    }
　}
}
</script>

<style>
@import url('https://fonts.googleapis.com/css?family=Source+Sans+Pro');

* {
  box-sizing: border-box;
  margin: 0;
  padding: 0;
}

body {
  font-family: 'Source Sans Pro', sans-serif;
}

#wrapper {
  background: radial-gradient(
    ellipse at top left,
    rgba(255, 255, 255, 1) 40%,
    rgba(229, 229, 229, 0.9) 100%
  );
  height: 100vh;
  padding: 60px 80px;
  width: 100vw;
}

#logo {
  height: auto;
  margin-bottom: 20px;
  width: 420px;
}

main {
  display: flex;
  justify-content: space-between;
}

main > div {
  flex-basis: 50%;
}

.left-side {
  display: flex;
  flex-direction: column;
}

.welcome {
  color: #555;
  font-size: 23px;
  margin-bottom: 10px;
}

.title {
  color: #2c3e50;
  font-size: 20px;
  font-weight: bold;
  margin-bottom: 6px;
}

.title.alt {
  font-size: 18px;
  margin-bottom: 10px;
}

.doc p {
  color: black;
  margin-bottom: 10px;
}

.doc button {
  font-size: 0.8em;
  cursor: pointer;
  outline: none;
  padding: 0.75em 2em;
  border-radius: 2em;
  display: inline-block;
  color: #fff;
  background-color: #4fc08d;
  transition: all 0.15s ease;
  box-sizing: border-box;
  border: 1px solid #4fc08d;
}

.doc button.alt {
  color: #42b983;
  background-color: transparent;
}
</style>
