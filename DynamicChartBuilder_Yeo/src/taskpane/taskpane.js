/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

const chartWidth = 500, chartHeight = 400;
//const dataSheetName = "Data";
// const dataTableName = "Table4";
const countryColumnName = "Countries";
const tempColumnName = "TempColumn";
const colorList = ["#afc97a","#cd7371","#729aca","#b65708","#276a7c","#4d3b62","#5f7530","#772c2a","#2c4d75","#f79646","#4bacc6","#8064a2","#9bbb59","#c0504d","#4f81bd"];
const fontSize = 20;
const chartName = "DynamicChart";

let logResult = document.getElementById("consoleText");
let columnCount = 0;

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    //document.getElementById("sideload-msg").style.display = "none";
    //document.getElementById("app-body").style.display = "flex";
    document.getElementById("createStaticChart").onclick = CreateFirstChart;
    document.getElementById("createDynamicChart").onclick = CreateDynamicChart;
  }
});

export async function CreateFirstChart() {
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
      columnCount = columns.count;
      let tempColumn = columns.getItemOrNullObject(tempColumnName);
      await context.sync();
      if (tempColumn.isNullObject) {
          tempColumn = columns.add(null, null, "TempColumn");
      }
      else {
          columnCount -= 1;
      }

      let tempRange = tempColumn.getDataBodyRange();
      let countryRange = columns.getItem(countryColumnName).getDataBodyRange();

      // Create Chart
      tempRange.copyFrom(columns.getItemAt(1).getDataBodyRange());
      let chart = dataSheet.charts.add(Excel.ChartType.barClustered, tempColumn.getRange());
      chart.set({ name: chartName, height: chartHeight, width: chartWidth });
      chart.title.format.font.set({size: fontSize});
      await context.sync();

      // Set category names
      let categoryAxis = chart.axes.getItem(Excel.ChartAxisType.category);
      categoryAxis.setCategoryNames(countryRange);
      categoryAxis.visible = false;
      let series = chart.series.getItemAt(0);
      series.set({ hasDataLabels: true, gapWidth: 30 });
      series.dataLabels.showCategoryName = true;
      series.points.load();
      await context.sync();
      writeLog(series.points.count);
      for (let i = 0; i < series.points.count; i++) {
          series.points.getItemAt(i).format.fill.setSolidColor(colorList[i % colorList.length]);
      }
      await context.sync();
    });
  } 
  catch (error) {
    writeLog(error);
  }

}

export async function CreateDynamicChart() {
  try {
    await Excel.run(async context => {
      const dataSheetName = document.getElementById("DataSheetName").value;
      const dataSheet = context.workbook.worksheets.getItem(dataSheetName);

      const dataTableName = document.getElementById("TableName").value;
      let table = dataSheet.tables.getItem(dataTableName);

      table.load("columns");
      await context.sync();
      let columns = table.columns;
      let tempColumn = columns.getItemOrNullObject(tempColumnName);
      let tempRange = tempColumn.getDataBodyRange();
      //let countryRange = columns.getItem(countryColumn).getDataBodyRange();
      tempRange.load("address");
      //countryRange.load("address");
      await context.sync();
      //console.log("column count: " + columnCount);
      writeLog("temp range:" + tempRange.address);
      //console.log("country range: " + countryRange.address);

      // let charts = dataSheet.charts;
      // let chart = charts.add(Excel.ChartType.barClustered, tempColumn.getRange());
      // chart.height = chartHeight;
      // chart.width = chartWidth;
      // chart.title.format.font.set({size: fontSize});
      // await context.sync();

      // let categoryAxis = chart.axes.getItem(Excel.ChartAxisType.category);
      // categoryAxis.setCategoryNames(countryRange);
      // let series = chart.series.getItemAt(0);
      // series.hasDataLabels = true;
      // series.gapWidth = 30;
      // series.dataLabels.showCategoryName = true;
      // series.points.load();
      // await context.sync();
      // console.log(countryRange);
      // console.log(series.points.count);
      // for (let i = 0; i < series.points.count; i++) {
      //     series.points.getItemAt(i).format.fill.setSolidColor(colorList[i % colorList.length]);
      // }
      // await context.sync();

      let chart = dataSheet.charts.getItem(chartName);

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
    writeLog(error);
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
      writeLog("count: " + columnCount);
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

function writeLog(log) {
  logResult.innerText = logResult.innerText + '\n' + log;
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


