# Lesson 4: Wrapping Up with Some New APIs

Some new functions were added recently and we can make use of them. First, make sure that you have the correct version of Office that supports the API requirement set.

## 4.0 Prep

4.0.1

For the ```calculate()``` function below, turn off the automatic formula calculation by going to File > Options > Formulas > select **Manual** under Workbook Calculation

To use the ```Range.Calculate()``` method and for ```calculate()```, see https://dev.office.com/reference/add-ins/excel/range

For the ConditionalFormat object, see:

- <https://github.com/OfficeDev/office-js-docs/blob/ExcelJs\_OpenSpec/reference/excel/conditionalformatcollection.md>

- <https://github.com/OfficeDev/office-js-docs/blob/ExcelJs\_OpenSpec/reference/excel/conditionalformat.md>

In order to use these two new APIs, you need Excel API requirement set 1.6 to be supported by the host application. Let's not take any chances and add a programmatic check in our code before executing these. Go here to read about requirement sets and how to check in code:
<https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets>

Now that you have the information, add a button to calculate our price table range B2:E7 (including our totals) and also to apply conditional formatting to the price table numbers in B3:E5.

4.1 Add the button to recalculate the range of the prices table.

4.2 In the same button handler, add code to apply conditional formatting.


Hints:

- It might be convenient to remove the stock multiply buttons from the sample code to make room for your new buttons.
