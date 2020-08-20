<template>
  <div id="app">
    <div>
      <div>
        <h2>常规用法</h2>
        <div>
          <table>
            <thead>
              <tr>
                <td>日期</td>
                <td>姓名</td>
                <td>省份</td>
                <td>市区</td>
                <td>地址</td>
              </tr>
            </thead>
            <tbody>
              <tr v-for="(item, i) in list" :key="i">
                <td v-for="(val, key) in item" :key="key">{{ val }}</td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>
    <FoxExport2excel :config="conf">导出excel</FoxExport2excel>
  </div>
</template>

<script>
import FoxExport2excel from "./components/json2excel";

export default {
  name: "App",
  components: {
    FoxExport2excel,
  },
  data() {
    return {
      list: [
        {
          date: "2016-05-03",
          name: "王小虎",
          province: "上海",
          city: "普陀区",
          address: "上海市普陀区金沙江路 1518 弄",
          zip: 200333,
        },
        {
          date: "2016-05-02",
          name: "王小虎",
          province: "上海",
          city: "普陀区",
          address: "上海市普陀区金沙江路 1518 弄",
          zip: 200333,
        },
        {
          date: "2016-05-04",
          name: "王小虎",
          province: "上海",
          city: "普陀区",
          address: "上海市普陀区金沙江路 1518 弄",
          zip: 200333,
        },
        {
          date: "2016-05-01",
          name: "王小虎",
          province: "上海",
          city: "普陀区",
          address: "上海市普陀区金沙江路 1518 弄",
          zip: 200333,
        },
        {
          date: "2016-05-08",
          name: "王小虎",
          province: "上海",
          city: "普陀区",
          address: "上海市普陀区金沙江路 1518 弄",
          zip: 200333,
        },
        {
          date: "2016-05-06",
          name: "王小虎",
          province: "上海",
          city: "普陀区",
          address: "上海市普陀区金沙江路 1518 弄",
          zip: 200333,
        },
        {
          date: "2016-05-07",
          name: "王小虎",
          province: "上海",
          city: "普陀区",
          address: "上海市普陀区金沙江路 1518 弄",
          zip: 200333,
        },
      ],
      conf: {
        multiHeader: [["id", "分类", "", "", "备注"]],
        header: ["", "土豆", "花生", "马铃薯", ""],
        merges: ["A1:A2", "B1:D1", "E1:E2"],
        data: [
          ["1", "2", "3", "4", "id对用的数量"],
          ["6", "7", "8", "9", "id对用的数量"],
        ],
      },
    };
  },
  mounted() {
    console.log(this.conf);
  },
};
</script>

<style lang="scss">
#app {
  font-family: Avenir, Helvetica, Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  text-align: center;
  color: #2c3e50;
  margin-top: 60px;
}
</style>
