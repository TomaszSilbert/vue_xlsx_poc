<template>
  <div id="app">
    <input type="file" @change="handleFile">
    
    <button v-on:click="exportWorkbook">XLSX Export</button>

    <div id="hitlist">
      <li v-for="item in stringRows" :key="item">{{ item }}</li>
    </div>
  </div>
</template>

<script lang="ts">
import { Component, Vue } from "vue-property-decorator";
import HomePage from "./components/HomePage.vue";
import XLSX from "xlsx";
import Template from "./Template";
import Template2 from "./Template2";
import { saveAs } from "file-saver";

@Component({
  components: {
    HomePage
  }
})
export default class App extends Vue {
  public stringRows: string[] = [];
  public workbook: XLSX.WorkBook | undefined = undefined;
  public template: Template | Template2 | undefined;

  handleFile(e) {
    this.stringRows = [];
    let context = this;
    let rABS = true; // true: readAsBinaryString ; false: readAsArrayBuffer
    let files = e.target.files,
      f = files[0];
    let reader = new FileReader();
    reader.onload = function(e) {
      let data = e.target.result;
      if (!rABS) data = new Uint8Array(data);
      context.workbook = XLSX.read(data, {
        type: rABS ? "binary" : "array"
      });
      let sheet: XLSX.WorkSheet = context.workbook.Sheets["Sheet1"];
      console.log(sheet);
      let range = XLSX.utils.decode_range(sheet["!ref"]);

      if (sheet["C1"].v === "Name") {
        context.template = new Template();
      }
      if (sheet["C1"].v === "Surname") {
        context.template = new Template2();
      }
      if (context.template) {
        for (let rowNum = range.s.r + 1; rowNum <= range.e.r; rowNum++) {
          let stringConc = "";
          let nameCell =
            sheet[
              XLSX.utils.encode_cell({ r: rowNum, c: context.template.name[0] })
            ];
          let surnameCell =
            sheet[
              XLSX.utils.encode_cell({
                r: rowNum,
                c: context.template.surname[0]
              })
            ];
          let emailCell =
            sheet[
              XLSX.utils.encode_cell({
                r: rowNum,
                c: context.template.email[0]
              })
            ];
          stringConc = stringConc
            .concat(nameCell.v)
            .concat(",")
            .concat(surnameCell.v)
            .concat(",")
            .concat(emailCell.v);
          context.stringRows.push(stringConc);
        }
      }
    };
    reader.readAsBinaryString(f);
  }
  exportWorkbook() {
    let template = this.template;
    //first of all let's update some cells
    // TODO
    if (this.workbook) {
      let sheet: XLSX.WorkSheet = this.workbook.Sheets["Sheet1"];
      if(template instanceof Template){
        if(sheet.H2){
          sheet.H2.v = "YES!";
        }else{
          sheet["H2"] = {a:"SheetJS", v:"Yes!"}
        }
        if(sheet.H3){
          sheet.H3.v = "YES!";
        }else{
          sheet["H3"] = {a:"SheetJS", v:"Yes!"}
        }
      }
    }

    if (this.workbook) {
      let wopts = { bookType: "xlsx", bookSST: false, type: "array" };
      let wbout = XLSX.write(this.workbook, wopts);
      // TODO check if there is a better way to save file
      saveAs(
        new Blob([wbout], { type: "application/octet-stream" }),
        "test.xlsx"
      );
    }
  }
}
</script>

<style>
#app {
  font-family: "Avenir", Helvetica, Arial, sans-serif;
}
#hitlist{
  padding: 30px;
}
</style>
