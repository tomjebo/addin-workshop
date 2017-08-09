# Lesson 2 Answers


2.0.1 Prep

```typescript
$("#run").click(run);
 
async function run() {
    try {
 
        await Excel.run(async (context) => {
            var range = context.workbook.getSelectedRange();
            range.format.fill.color = "yellow";
            range.load(["address", "values"]);
            await context.sync()
            console.log("The range address was \"" + range.address + "\".");
            await populateRange(context, range);
        })
    }
     catch (error) {
            OfficeHelpers.UI.notify(error);
            OfficeHelpers.Utilities.log(error);
        }
}

async function populateRange(context: Excel.RequestContext, range: Excel.Range) {
    console.log("populateRange: range is - ", range.address);
    var newValues = range.values;
    var counter = 1;
    for (var i = 0; i < newValues.length; i++) {
        for (var j = 0; j < newValues[i].length; j++) {
            newValues[i][j] = counter++;
        }
    }
    range.values = newValues;
 
    await context.sync()
            console.log("finished populating the matrix");
 
}

```

2.5 Grand Total button

```html
// Added this block of code
<button id="grand-total" class="ms-Button">
        <span class="ms-Button-label">Grand Total</span>
</button>
```

```typescript
async function grandTotal() {
    try {
        await Excel.run(async (ctx) => {
            var range = ctx.workbook.worksheets.getItem("Sample").getRange("E3:E5");
            var rangeTot = ctx.workbook.worksheets.getItem("Sample").getRange("B7:E8");
            var gTot = ctx.workbook.functions.sum(range);

            range.load("values");
            rangeTot.load("values");
            gTot.load();

            await ctx.sync();

            var vTot = rangeTot.values;

            console.log(gTot.value);
            console.log(range);
            vTot[0][3] = gTot.value;
            vTot[0][0] = "Grand Total";
            vTot[0][1] = "=sum(c3:c5)";

            rangeTot.values = vTot;

            await ctx.sync();
        });
    }
    catch (error) {
        OfficeHelpers.UI.notify(error);
        OfficeHelpers.Utilities.log(error);
    }
}
```
