<template>
    <div id="app">
        <div class="ctx">
            <div>
                <h2>常规用法</h2>
                <div>
                    <div>
                        <FoxExport2excel :config="conf">导出excel</FoxExport2excel>
                    </div>
                    <table>
                        <thead>
                            <tr>
                                <td>日期</td>
                                <td>姓名</td>
                                <td>省份</td>
                                <td>市区</td>
                                <td>地址</td>
                                <td>邮编</td>
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
            <div>
                <h2>多级表头</h2>
                <div>
                    <div>
                        <FoxExport2excel :config="conf">导出excel</FoxExport2excel>
                    </div>
                    <table>
                        <thead>
                            <tr>
                                <td rowspan="3">日期</td>
                                <td colspan="5">配送信息</td>
                            </tr>
                            <tr>
                                <td rowspan="2">姓名</td>
                                <td colspan="4">地址</td>
                            </tr>
                            <tr>
                                <td>省份</td>
                                <td>市区</td>
                                <td>地址</td>
                                <td>邮编</td>
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
body {
    padding: 0;
    margin: 0;
}
#app {
    color: #333;
    height: 100vh;
    width: 100vw;
    overflow: auto;
    font-size: 14px;
}
table {
    border-collapse: collapse;
}
thead {
    background-color: #f8f8f8;
}
td {
    border: 1px solid #dfdfdf;
    padding: 4px 10px;
}

.ctx {
    width: 800px;
    margin: 0 auto;
}
</style>
