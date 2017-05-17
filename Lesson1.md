# Lesson 1: Warm Up


## Setup

1.1 Open Excel (preferably the desktop version).

1.2 Go to Insert > My Add-ins > the Store icon (red shopping bag).

1.3 Search for "Script Lab".

1.4 Click **Add**.

1.5 Click on the **Script Lab** tab and see the  **Code** and **Run** commands.


## Running the First Sample

1.6 In the **Code** pane, select **Samples**.

1.7 Select **Basic API** call (JavaScript).

1.8 In the **Run** pane, select the same.

1.9 Select a matrix of several cells and click the **Run Code** button.

Observations:

-   The cells selected should be highlighted in yellow.
-   Review the code in the Code pane.

Notice the ```Excel.run()``` invocation
```

$("#run").click(run);

function run() {
    Excel.run(function (context) {
        var range = context.workbook.getSelectedRange();
        range.format.fill.color = "yellow";
        range.load("address");
        return context.sync()
            .then(function() {
                console.log("The range address was \"" + range.address + "\".");
            });
    })
        .catch(function(error) {
            OfficeHelpers.UI.notify(error);
            OfficeHelpers.Utilities.log(error);
        });
}
```
Note the ```context.sync()``` and ```.then``` pattern. Asynchronous code must always
return a Promise.

Notice the output from the ```console.log()``` call on line 10 on the Firebug
Console tab.

Adding some functionality:

You can now edit the Basic API call (JavaScript) sample code and it will save it to your snippets.

1.10 Using the Script Lab code editor, modify the code to populate the cells with increasing numbers starting at 1.

For example, if four cells are selected,

![alt text](Image1_lesson1.png)

Hints:

-   Use the ```.values``` property of the Range object.
-   Remember to load "values" first and then sync.
-   Fewer calls to ```context.sync``` mean fewer calls to the Office application.
-   If you use another function to populate, remember to pass in the context as well as the range.
-   If you get stuck, look at the other sample snippets for ideas.

1.11 Once satisfied, run it to show the populated cells.

1.12 Now you have a modified version of the Basic API call (JavaScript) sample code in MySnippets. Navigate to see that it's there.

This modified code in MySnippets will only remain in the add-in memory until you clear your browser cache. We'll discuss saving and sharing the code in another lesson.
