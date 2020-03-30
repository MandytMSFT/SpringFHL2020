(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;

    const chartWidth = 500, chartHeight = 400;
    //const dataSheetName = "Data";
   // const dataTableName = "Table4";
    const countryColumn = "Countries";
    const tempColumnName = "TempColumn";
    const colorList = ["red", "green", "blue", "grey", "yellow", "brown","purple"];

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the FabricUI notification mechanism and hide it
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();
            
            // If not using Excel 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
                $("#template-description").text("This sample will display the value of the cells that you have selected in the spreadsheet.");
                $('#button-text').text("Display!");
                $('#button-desc').text("Display the selection");

                $('#highlight-button').click(displaySelectedCells);
                return;
            }

          //  $("#template-description").text("This sample highlights the highest value from the cells you have selected in the spreadsheet.");
            $('#button-text').text("CreateDynamicChart");
           // $('#button-desc').text("Highlights the largest number.");
                
          //  loadSampleData();

            // Add a click event handler for the highlight button.
            $('#highlight-button').click(() => tryCatch(CreateDynamicChart));
        });
    };

    function CreateDynamicChart() {
        Excel.run(async (context) => {
            const dataSheetName = document.getElementById("DataSheetName").value;
            const dataSheet = context.workbook.worksheets.getItem(dataSheetName);

            const dataTableName = document.getElementById("TableName").value;
            let table = dataSheet.tables.getItem(dataTableName);

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
            await context.sync();
            let categoryAxis = chart.axes.getItem(Excel.ChartAxisType.category);
            categoryAxis.setCategoryNames(countryRange);
            let series = chart.series.getItemAt(0);
            series.hasDataLabels = true;
            series.points.load();
            await context.sync();
            console.log(countryRange);
            console.log(series.points.count);
            for (let i = 0; i < series.points.count; i++) {
                series.points.getItemAt(i).format.fill.setSolidColor(colorList[i % colorList.length]);
            }
            await context.sync();
            //for (let i = columnCount - 1; i > 0; i--) {
            for (let i = 1; i < columnCount; i++) {
                let dataRange = columns.getItemAt(i).getDataBodyRange();
                let titleRange = columns.getItemAt(i).getHeaderRowRange();
                titleRange.load("text");
                await context.sync();
                //series.hasDataLabels = true;
                //chart.title.set({ text: titleRange.text[0][0]});
                chart.title.text = titleRange.text[0][0];
                tempRange.copyFrom(dataRange);
                table.sort.apply([{ key: columnCount, ascending: true }], true);
                await context.sync();
                sleep(300);
            }
            //tempColumn.delete();
            await context.sync();
        });
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
})();
