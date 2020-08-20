<template>
  <component :is="tag" class="hello" @click="handleDownload"
    ><slot></slot
  ></component>
</template>

<script>
import { export_json_to_excel } from "@/utils/json2excel.js";
export default {
  name: "foxExport2excel",
  props: {
    config: {
      type: Object,
      default() {
        return {};
      },
    },
    tag: {
      type: String,
      default: "button",
    },
  },
  methods: {
    handleDownload() {
      const conf = this.config;
      export_json_to_excel(conf);
    },
  },
};
</script>

<!-- Add "scoped" attribute to limit CSS to this component only -->
<style scoped lang="scss">
h3 {
  margin: 40px 0 0;
}
ul {
  list-style-type: none;
  padding: 0;
}
li {
  display: inline-block;
  margin: 0 10px;
}
a {
  color: #42b983;
}
</style>
