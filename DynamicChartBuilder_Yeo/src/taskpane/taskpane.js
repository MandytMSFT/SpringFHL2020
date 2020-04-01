/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

const chartWidth = 500, chartHeight = 400;
//const dataSheetName = "Data";
// const dataTableName = "Table4";
const countryColumn = "Countries";
const tempColumnName = "TempColumn";
const colorList = ["#afc97a","#cd7371","#729aca","#b65708","#276a7c","#4d3b62","#5f7530","#772c2a","#2c4d75","#f79646","#4bacc6","#8064a2","#9bbb59","#c0504d","#4f81bd"];
const fontSize = 20;

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    //document.getElementById("sideload-msg").style.display = "none";
    //document.getElementById("app-body").style.display = "flex";
    document.getElementById("createDynamicChart").onclick = CreateDynamicChart;
  }
});

export async function CreateDynamicChart() {
  try {
    await Excel.run(async context => {
      // Find selected table
      let activeRange = context.workbook.getSelectedRange();
      let dataTables = activeRange.getTables(false);
      dataTables.load("items");
      await context.sync();

      // Get active table
      let dataTable = dataTables.items[0];
      let dataSheet = context.workbook.worksheets.getActiveWorksheet();
      let table = dataSheet.tables.getItem(dataTable.id);

      table.load("columns");
      await context.sync();
      let columns = table.columns;
      let columnCount = columns.count;
      let tempColumn = columns.getItemOrNullObject(tempColumnName);
      await context.sync();
      if (tempColumn.isNullObject) {
          tempColumn = columns.add(null, null, "TempColumn");
      }
      else {
          columnCount -= 1;
      }
      let tempRange = tempColumn.getDataBodyRange();
      let countryRange = columns.getItem(countryColumn).getDataBodyRange();
      tempRange.load("address");
      countryRange.load("address");
      await context.sync();
      console.log("column count: " + columnCount);
      console.log("temp range:" + tempRange.address);
      console.log("country range: " + countryRange.address);
      let charts = dataSheet.charts;
      let chart = charts.add(Excel.ChartType.barClustered, tempColumn.getRange());
      chart.height = chartHeight;
      chart.width = chartWidth;
      chart.title.format.font.set({size: fontSize});
      await context.sync();

      let categoryAxis = chart.axes.getItem(Excel.ChartAxisType.category);
      categoryAxis.setCategoryNames(countryRange);
      let series = chart.series.getItemAt(0);
      series.hasDataLabels = true;
      series.gapWidth = 30;
      series.dataLabels.showCategoryName = true;
      series.points.load();
      await context.sync();
      console.log(countryRange);
      console.log(series.points.count);
      for (let i = 0; i < series.points.count; i++) {
          series.points.getItemAt(i).format.fill.setSolidColor(colorList[i % colorList.length]);
      }
      await context.sync();

      let interval = document.getElementById("ChartSpeed").value
      //for (let i = columnCount - 1; i > 0; i--) {
      for (let i = 1; i < columnCount; i++) {
          let dataRange = columns.getItemAt(i).getDataBodyRange();
          let titleRange = columns.getItemAt(i).getHeaderRowRange();
          titleRange.load("text");
          await context.sync();
          //chart.title.set({ text: titleRange.text[0][0]});
          chart.title.text = titleRange.text[0][0];
          tempRange.copyFrom(dataRange);
          table.sort.apply([{ key: columnCount, ascending: true }], true);
          await context.sync();

          sleep(interval);
      }
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

// Pivot table cannot sort and data point color doesn't keep the same when change data hierachy.
async function testPivotTable() {
  await Excel.run(async (context) => {
      let book = context.workbook;
      let sheet = context.workbook.worksheets.getItem(dataSheetName);
      let pivotSheet = book.worksheets.getItem(pivotSheetName);
      // Clear pivot sheet
      let deprecatedRange = pivotSheet.getRange(null);
      deprecatedRange.clear();
      await context.sync();
      let table = sheet.tables.getItemAt(0);
      let pivotTable = book.pivotTables.add("PivotTable1", table.getRange(), pivotSheet.getRange("A1"));
      //await context.sync();
      pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem(countryColumn));
      pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("2020-03-25"));
      await context.sync();
  });
}
async function testTable() {
  await Excel.run(async (context) => {
      let book = context.workbook;
      let sheet = context.workbook.worksheets.getItem(dataSheetName);
      let table = sheet.tables.getItemAt(0);
      table.load("columns");
      await context.sync();
      let columns = table.columns;
      let columnCount = columns.count;
      console.log("count: " + columnCount);
      table.sort.apply([{ key: 1, ascending: true }], true);
      await context.sync();
  });
}
function sleep(sleepTime) {
  var start = new Date().getTime();
  for (var i = 0; i < 1e7; i++) {
      if (new Date().getTime() - start > sleepTime) {
          break;
      }
  }
}
/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
      await callback();
  } catch (error) {
      // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
      console.error(error);
  }
}


