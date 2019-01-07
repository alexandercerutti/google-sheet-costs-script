# Costs Script for Google Spreadsheets

I was in need for a semi-automated way to track my per-month costs. I discovered Google Apps APIs and I created this script that automatically creates a sheet per month with the following columns:
 - Expense name
 - Cost
 - Date
 - Notes

![Header and expenses](assets/expenses.png)

Each inserted record will auto-update the total of the costs for that month, by using a `=SUM()` function on the costs column, with a range from its second row (see first screenshow), to the first gray row (see second screenshot) in the footer. By doing this, the references will update when a new row is created and you won't have to update manually the formula.

When multiple month-sheets (like `October 2018`, `November 2018`, etc.) will be available, every sheet with "index" greater than 0 (e.g. I started from October (0) and now we are in November (1)), will have a total row to keep track of the amount of money (positive or negative) left based on the previous months.

E.g in October I spent 100€ and got an income of 200 €. I'll be in positive of 100€ (shown as -100, since we are talking about expenses, not earnings).
In November I spent 150€ but still not have received any income yet. So I'll have 100€ left from Oct. minus 150€ spent this month. I will get in the last rows 150€ and 50€ (last one highlighted in red but expressed in positive, again because we are talking about expenses, not earnings).

![Footer](assets/footer.png)

### Usage

To use it you have to create first a new Google Drive SpreadSheet document. In the toolbar go to `Tools > Script editor`. A new tab will open.
Create a new script and paste `index.js` script in it.

At the beginning of the script, some settings are available to be customized.


| Property | Description | Type | Default Value |
|----------|-------------|------|---------------|
| `trackMonthly` | To create a new sheet every month. | Boolean | `true` |
| `pastMonthTotal` | To count the previous month costs. | Boolean | `true` |
| `excludeSheets` | To exclude some sheets to be edited on last row editing. | Array\<String> | `[]` |
| `columnsDefaultValue` | To set a new sheet columns default values. As you see you can include a function to be executed with binded parameters. Use null to exclude a column from default value. | Array\<String> | `[null, null, parsedDate.bind(this, false), "-"]` |
| `columnsName` | To set a new sheet columns header names. | Array\<String> | `["Nome", "Costo", "Data", "Note"]` |
| `columnsWidth` | To set new sheet columns width. | Array\<Number> | `[185, 185, 185, 200]` |
| `rowsHeight` | To set new sheet default rows height. | Number | `35` |


For any further explanation, open a topic in issues. Default language is Italian.

### Development

To work on any improvements of this script, I suggest you to install GAS Typings (Typescript):

```sh
$ npm install --save @types/google-apps-script
```

So you will have the latest API reference directly in your editor (better if VSCode or one that supports `DefinitelyTyped`).

Hope you enjoy! 😊
