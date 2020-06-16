<template>
  <div id="app">
    <a
      href="https://github.com/seungwoo321/vue-pivottable"
      target="_blank"
      style="position: fixed; top: 1rem; right: 1rem;"
    >
      <svg id="i-github" viewBox="0 0 64 64" width="36" height="36">
        <path
          stroke-width="0"
          fill="black"
          d="M32 0 C14 0 0 14 0 32 0 53 19 62 22 62 24 62 24 61 24 60 L24 55 C17 57 14 53 13 50 13 50 13 49 11 47 10 46 6 44 10 44 13 44 15 48 15 48 18 52 22 51 24 50 24 48 26 46 26 46 18 45 12 42 12 31 12 27 13 24 15 22 15 22 13 18 15 13 15 13 20 13 24 17 27 15 37 15 40 17 44 13 49 13 49 13 51 20 49 22 49 22 51 24 52 27 52 31 52 42 45 45 38 46 39 47 40 49 40 52 L40 60 C40 61 40 62 42 62 45 62 64 53 64 32 64 14 50 0 32 0 Z"
        />
      </svg>
    </a>
    <div class="title">
      <h1>Vue Pivottable</h1>
      <small>Sample Dataset: Tips</small>
    </div>
    <vue-pivottable-ui
      :data="pivotData"
      :aggregatorName="aggregatorName"
      :rendererName="rendererName"
      :rows="rows"
      :cols="cols"
      :vals="vals"
      :disabledFromDragDrop="disabledFromDragDrop"
      :sortonlyFromDragDrop="sortonlyFromDragDrop"
      :hiddenFromDragDrop="hiddenFromDragDrop"
    ></vue-pivottable-ui>

<br/>
<button>csv</button>
    <footer>
      Released under the
      <a href="//github.com/seungwoo321/vue-pivottable/blob/master/LICENSE">MIT</a> license.
      <a href="//github.com/seungwoo321/vue-pivottable">View source.</a>
    </footer>
  </div>
</template>
<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.3.1/jquery.min.js"></script>

<script>
import tips from './tips'
import JQuery from 'jquery'
window.$ = JQuery
// import { VuePivottable, VuePivottableUi } from 'vue-pivottable'
import { VuePivottableUi } from 'vue-pivottable'
import 'vue-pivottable/dist/vue-pivottable.css'
export default {
  components: {
    VuePivottableUi,
  },
  name: 'app',
  data () {
    return {
      pivotData: tips,
      aggregatorName: 'Sum',
      rendererName: 'Table',
      rows: ['Payer Gender', 'Party Size'],
      cols: ['Meal', 'Payer Smoker', 'Day of Week'],
      vals: ['Total Bill'],
      disabledFromDragDrop: ['Payer Gender'],
      hiddenFromDragDrop: ['Total Bill'],
      sortonlyFromDragDrop: ['Party Size']
    }
  },
  methods: {},
  mounted() {
    $("button").click(function() {
          exportTableToExcel();
        });
  }
  }
  function write_headers_to_excel(table) {
      var str="";
      console.log(table)
      var myTableHead = table.getElementsByTagName('thead')[0];
      var rowCount = myTableHead.rows.length;
      var colCount = myTableHead.getElementsByTagName("tr")[0].getElementsByTagName("th").length;

      var ExcelApp = new ActiveXObject("Excel.Application");
      var ExcelSheet = new ActiveXObject("Excel.Sheet");
      ExcelSheet.Application.Visible = true;

      for(var i=0; i<rowCount; i++) {
        for(var j=0; j<colCount; j++) {
          str= myTableHead.getElementsByTagName("tr")[i].getElementsByTagName("th")[j].innerHTML;
          ExcelSheet.ActiveSheet.Cells(i+1,j+1).Value = str;
          }
      }
  }

  function write_bodies_to_excel (table) {
      var str="";

      var myTableHead = table.getElementsByTagName('tbody')[0];
      var rowCount = myTableHead.rows.length;
      var colCount = myTableHead.getElementsByTagName("tr")[0].getElementsByTagName("th").length;

      var ExcelApp = new ActiveXObject("Excel.Application");
      var ExcelSheet = new ActiveXObject("Excel.Sheet");
      ExcelSheet.Application.Visible = true;

      for(var i=0; i<rowCount; i++) {
        for(var j=0; j<colCount; j++) {
          str= myTableHead.getElementsByTagName("tr")[i].getElementsByTagName("th")[j].innerHTML;
          ExcelSheet.ActiveSheet.Cells(i+1,j+1).Value = str;
          }
      }
  }

  function exportTableToExcel(tableID, filename = ''){
    var downloadLink;
    var dataType = 'application/vnd.ms-excel';
    $('table.pvtTable').attr('border', '1');
    var tableSelect = document.getElementsByClassName('pvtTable');
    console.log(tableSelect)
    var tableHTML = tableSelect[0].outerHTML.replace(/ /g, '%20');

    // Specify file name
    filename = filename?filename+'.xls':'excel_data.xls';

    // Create download link element
    downloadLink = document.createElement("a");

    document.body.appendChild(downloadLink);

    if(navigator.msSaveOrOpenBlob){
        var blob = new Blob(['\ufeff', tableHTML], {
            type: dataType
        });
        navigator.msSaveOrOpenBlob( blob, filename);
    }else{
        // Create a link to the file
        downloadLink.href = 'data:' + dataType + ', ' + tableHTML;

        // Setting the file name
        downloadLink.download = filename;
        //triggering the function
        downloadLink.click();
    }
}
</script>

<style>
.main {
  max-width: 980px;
  margin: 8vh auto 20px;
}
.title {
  text-align: center;
  margin-bottom: 20px;
}
h1 {
  margin-bottom: 0px;
}
.table-responsive {
  display: block;
  width: 100%;
  overflow-x: auto;
}
pre {
  text-align: left;
  background-color: #f8f8f8;
  padding: 1.2em 1.4em;
  line-height: 1.5em;
  margin: 60px 0 0;
  overflow: auto;
}
code {
  padding: 0;
  margin: 0;
}
footer {
  text-align: center;
  margin-top: 40px;
  line-height: 2;
}
</style>
