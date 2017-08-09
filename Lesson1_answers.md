
# Lesson 1 Answers

1.10 Modified code to populate the range

```javascript
$("#run").click(run);

function run() {
    Excel.run(function (context) {
        var range = context.workbook.getSelectedRange();
        range.format.fill.color = "yellow";
        range.load([ "address", "values"]);
        return context.sync()
            .then(function () {
                console.log("The range address was \"" + range.address + "\".");
                return populateRange(context, range); // Added this line of code

            });

    })
        .catch(function (error) {
            OfficeHelpers.UI.notify(error);
            OfficeHelpers.Utilities.log(error);
        });
}

// Added the following code block
function populateRange(context: Excel.RequestContext, range: Excel.Range) {
    console.log("populateRange: range is - ", range.address);
            var newValues = range.values;
            var counter = 1;
            for (var i = 0; i < newValues.length; i++) {
                for (var j = 0; j < newValues[i].length; j++) {
                    newValues[i][j] = counter++;
                }
            }
            range.values = newValues;

            return context.sync()
	    .then(function () {
			    console.log("finished populating the matrix");
			    });
}
```