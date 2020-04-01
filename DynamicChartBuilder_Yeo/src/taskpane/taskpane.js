/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// Parameters. Modify it if needed.
const chartWidth = 500, chartHeight = 400;
const splitIncreasement = 3;
const colorList = ["#afc97a","#cd7371","#729aca","#b65708","#276a7c","#4d3b62","#5f7530","#772c2a","#2c4d75","#f79646","#4bacc6","#8064a2","#9bbb59","#c0504d","#4f81bd"];
const fontSize = 20;

// Internal used const. DO NOT CHANGE
const tempColumnName = "TempColumn";
const increaseColumnName = "IncreaseColumn";
const colorColumnName = "colorColumn";
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

      // Get columns
      let columns = table.columns;
      columnCount = columns.count;
      let tempColumn = columns.getItemOrNullObject(tempColumnName);
      let increaseColumn = columns.getItemOrNullObject(increaseColumnName);
      let colorColumn = columns.getItemOrNullObject(colorColumnName);
      await context.sync();
      if (tempColumn.isNullObject) {
          tempColumn = columns.add(null, null, "TempColumn");
      }
      else {
          columnCount -= 1;
      }
      if (increaseColumn.isNullObject) {
        increaseColumn = columns.add(null, null, increaseColumnName);
      } else {
        columnCount -= 1;
      }
      if (colorColumn.isNullObject) {
        colorColumn = columns.add(null, null, colorColumnName);
      } else {
        columnCount -= 1;
      }

      // Get ranges
      let countryRange = columns.getItemAt(0).getDataBodyRange();
      let tempRange = tempColumn.getDataBodyRange();
      let increaseRange = increaseColumn.getDataBodyRange();
      let colorRange = colorColumn.getDataBodyRange();
      tempRange.clear();
      increaseRange.clear();
      colorRange.load("values");
      await context.sync();

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
      categoryAxis.set({ visible: true });
      let series = chart.series.getItemAt(0);
      series.set({ hasDataLabels: true, gapWidth: 30 });
      series.dataLabels.showCategoryName = false;
      series.points.load();
      await context.sync();
      //writeLog(series.points.count);

      // Set data points color
      for (let i = 0; i < series.points.count; i++) {
          colorRange.getCell(i, 0).values = colorList[i % colorList.length];
          series.points.getItemAt(i).format.fill.setSolidColor(colorList[i % colorList.length]);
          writeLog(colorList[i % colorList.length]);
      }
      //series.points.load();
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
      let increaseColumn = columns.getItemOrNullObject(increaseColumnName);
      let colorColumn = columns.getItemOrNullObject(colorColumnName);
      let tempRange = tempColumn.getDataBodyRange();
      let increaseRange = increaseColumn.getDataBodyRange();
      let colorRange = colorColumn.getDataBodyRange();
      //tempRange.load("address");
      await context.sync();
      //writeLog("temp range:" + tempRange.address);

      let chart = dataSheet.charts.getItem(chartName);
      let interval = document.getElementById("ChartSpeed").value;

      for (let i = 1; i < columnCount; i++) {
          let dataRange = columns.getItemAt(i).getDataBodyRange();
          tempRange.load("values");
          dataRange.load("values");
          increaseRange.clear();
          await context.sync();

          // Calculate increase based on current value and next value
          let increaseData = calculateIncrease(tempRange.values, dataRange.values, splitIncreasement);
          for (let k = 0; k < increaseData.length; k ++) {
            increaseRange.getCell(k, 0).values = increaseData[k] | 0;
          }
          increaseRange.setDirty();
          increaseRange.calculate();
          await context.sync();
    


          for (let j = 1; j <= splitIncreasement; j++) {
            if (j == splitIncreasement) {
              // Directly use next column data
              tempRange.copyFrom(dataRange);
            }
            else {
              // Add increase amount
              tempRange.load("values");
              increaseRange.load("values");
              await context.sync();
              for (let k = 0; k < tempRange.values.length; k++) {
                tempRange.getCell(k, 0).values = tempRange.values[k][0] + increaseRange.values[k][0];
              }
            }

            tempRange.setDirty();
            tempRange.calculate();
            //await context.sync();
            table.sort.apply([{ key: columnCount, ascending: true }], true);

            // Set data points color
            let series = chart.series.getItemAt(0);
            series.load("points");
            colorRange.load("values");
            await context.sync();
            for (let k = 0; k < series.points.count; k++) {
              series.points.getItemAt(k).format.fill.setSolidColor(colorRange.values[k][0]);
            }

            //console.log("Current Value:" + tempRange.values);
            await context.sync();
            tempRange.load("values");
            await context.sync();
            sleep(interval);
          }

          
          let titleRange = columns.getItemAt(i).getHeaderRowRange();
          titleRange.load("text");
          await context.sync();

          chart.title.text = titleRange.text[0][0];
          await context.sync();
      }

      colorRange.load("values");
      series.load("points");
      await context.sync();

      // Set color again after all done
      for (let k = 0; k < series.points.count; k++) {
        series.points.getItemAt(k).format.fill.setSolidColor(colorRange.values[k][0]);
      }
      await context.sync();
    });
  } catch (error) {
    writeLog(error);
  }
}

// To calculate the increase for each step between next data list and current data list
//function calculateIncrease(current: Array<Array<number>>, next: Array<Array<number>>, steps: number) {
function calculateIncrease(current, next, steps) {
  if (current.length != next.length) {
    console.error("Error! current data length:" + current.length + ", next data length" + next.length + ".");
  }

  let result = new Array(current.length);
  for (let i = 0; i < current.length; i++) {
    let increasement = (next[i][0] - current[i][0]) / steps;
    result[i] = increasement;
  }

  return result;
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


