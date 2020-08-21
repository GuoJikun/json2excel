import Json2excel from "./src/json2excel.vue";

Json2excel.install = Vue => {
    Vue.component(Json2excel.name, Json2excel);
};

export default Json2excel;
