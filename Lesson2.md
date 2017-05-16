# Lesson 2: TypeScript


## Prep for TypeScript

Up to this point, you've used JavaScript to call the APIs only. You haven't
tried TypeScript yet.

To prepare for this lesson, do a simple modification to the code from Lesson 1.

2.0.1 Convert the return context.sync .then construct to the simpler and more readable async/await TypeScript construct.

Hints:

- A function must be declared with the ```async``` modifier if ```await``` is called inside.

- ```await``` replaces the return and ```.then``` wrapper around the completion code.

- All TypeScript types and libraries have already been included, check them out on the Libraries tab.

2.0.2 Make sure that the code runs successfully as it did in Lesson 1.

Now let's switch gears. Start with a new sample snippet and this time add some more realistic functionality.

## Setup

2.1 Navigate and open the sample called "Copy and multiply values".

Observation: You can see that this is a TypeScript example. Note the await/async pattern and how much more readable this is.

2.2 Refresh the content in the Run pane and click the **Add sample data** button to see data inserted into the sheet.

2.3 Switch to the Template tab to see the HTML that drives the UI.

Observation: "Add sample data" is a button element with id="setup" and this is hooked up with the click handler called setup().

2.4 Look at the HTML code and study what each button's function is. Back in the **Run** pane, try the one labeled, "Multiply values using for loop". Pretty straight forward, right?

Now add a button and handler to create a Grand Total under the Total Price column. For this, use the sum formula.

2.5 Add another button with a label of "Grand Total".

2.6 Add code to total the Total Price column and put the result in E7 (below the last entry). Also add the label "Grand Total" in B7.

Hints:

- Use https://dev.office.com/reference/add-ins/excel/functions
- Remember that the values array will index from 0, even though the Excel
addresses are 1 based.
- Use the Workbook.functions.sum() method.

Notice that the sum() method returns programmatically the value of the sum of the range, which we add to a cell. However, typically, you'd add a formula like this: ``` "=sum(<range>)" ``` into that cell instead of the resulting value.

2.7 Use the range calculation API in Lesson 4, so let's add another value into the Grand Total row. This one should total up the Qty column but not use the workbook.function.sum() method. Instead add the "=sum()" formula into the cell for later calculation.

2.8 After this is successful, add another row with Tax (say B6:E6) and include that into the Grand Total amount.



