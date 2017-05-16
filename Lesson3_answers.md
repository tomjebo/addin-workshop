# Lesson 3 Answers

```
async function createChart() {
    try {
        await Excel.run(async (ctx) => {
            var rangeSelection = "B2:E5";
            var range = ctx.workbook.worksheets.getItem("Sample")
                .getRange(rangeSelection);
            var chart = ctx.workbook.worksheets.getItem("Sample")
                .charts.add("ColumnClustered", range, "auto");
            await ctx.sync();
            console.log("New Chart Added");
        });
    }
    catch (error) {
        OfficeHelpers.UI.notify(error);
        OfficeHelpers.Utilities.log(error);
    }
}
```