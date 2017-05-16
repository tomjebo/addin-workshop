# Lesson 4 Answers

4.0.1 Programmatically check for API set 1.6.

```
if (Office.context.requirements.isSetSupported('ExcelApi', 1.6) === true) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

4.1 Recalculate!

```
async function recalculate() {
    try {
        await Excel.run(async (ctx) => {
            console.log("Recalculating price table");
            if (Office.context.requirements.isSetSupported('ExcelApi', 1.6) === true) {
                var rangeSelection = "B2:E5";
                var range = ctx.workbook.worksheets.getItem("Sample")
                    .getRange(rangeSelection);
                range.calculate();
                await ctx.sync();
                console.log("Done recalculating price table!");
            }
            else {
                console.log("Can't recalculate in this host");
            }

        });
    }
    catch (error) {
        OfficeHelpers.UI.notify(error);
        OfficeHelpers.Utilities.log(error);
    }
}
```

4.2 Add conditional formatting.

```
async function recalculate() {
	try {
		await Excel.run(async (ctx) => {
				console.log("Recalculating price table");
				if (Office.context.requirements.isSetSupported('ExcelApi', 1.6) === true) {
				var rangeSelection = "B2:E5";
				var range = ctx.workbook.worksheets.getItem("Sample")
				.getRange(rangeSelection);
				range.calculate();
				await ctx.sync();

				var conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
				conditionalFormat.iconOrNull.style = "YellowThreeArrows";
				await ctx.sync()
				console.log("Added new yellow three arrow icon set.");
				console.log("Done recalculating price table!");
				}
				else {
				console.log("Can't recalculate in this host");
				}

				});
	}
	catch (error) {
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
		OfficeHelpers.UI.notify(error);
		OfficeHelpers.Utilities.log(error);
	}
}
```