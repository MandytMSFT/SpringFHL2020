/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

const chartWidth = 500, chartHeight = 400;
const countryColumnName = "Countries";
const tempColumnName = "TempColumn";
const colorList = ["#afc97a","#cd7371","#729aca","#b65708","#276a7c","#4d3b62","#5f7530","#772c2a","#2c4d75","#f79646","#4bacc6","#8064a2","#9bbb59","#c0504d","#4f81bd"];
const fontSize = 20;
const chartName = "DynamicChart";

let logResult = document.getElementById("consoleText");
let columnCount = 0;
let activeTableId;

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("createFirstChart").onclick = CreateFirstChart;
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
      activeTableId = dataTable.id;
      let table = dataSheet.tables.getItem(activeTableId);

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
      let dataColumn = columns.getItemAt(1); // Use the first data column as starting chart data
      tempRange.copyFrom(dataColumn.getDataBodyRange());
      table.sort.apply([{ key: columnCount, ascending: true }], true);
      let chart = dataSheet.charts.add(Excel.ChartType.barClustered, tempColumn.getRange());
      chart.set({ name: chartName, height: chartHeight, width: chartWidth });
      let headerRange = dataColumn.getHeaderRowRange();
      headerRange.load("text");
      await context.sync();

      // Set chart tile and style
      chart.title.text = headerRange.text[0][0];
      chart.title.format.font.set({size: fontSize});
      chart.legend.set({ visible: false });

      // Set category names
      let categoryAxis = chart.axes.getItem(Excel.ChartAxisType.category);
      categoryAxis.setCategoryNames(countryRange);
      categoryAxis.visible = false;
      let series = chart.series.getItemAt(0);
      series.set({ hasDataLabels: true, gapWidth: 30 });
      series.dataLabels.showCategoryName = true;
      series.points.load();
      await context.sync();
      //writeLog(series.points.count);

      // Set data points color
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
      let dataSheet = context.workbook.worksheets.getActiveWorksheet();
      let table = dataSheet.tables.getItem(activeTableId);

      table.load("columns");
      await context.sync();
      let columns = table.columns;
      let tempColumn = columns.getItemOrNullObject(tempColumnName);
      let tempRange = tempColumn.getDataBodyRange();
      //tempRange.load("address");
      await context.sync();
      //writeLog("temp range:" + tempRange.address);

      let chart = dataSheet.charts.getItem(chartName);
      let interval = document.getElementById("ChartSpeed").value;

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


