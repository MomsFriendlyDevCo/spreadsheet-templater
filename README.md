Spreadsheet-Handlebars
======================
[Handlebars](http://handlebarsjs.com) integration for spreadhssets (via [XLSX](https://docs.sheetjs.com)).

This plugin allows a spreadsheet to use handlebars notation to replace cell contents which enables an input spreadsheet to act as a template for incomming data.


```javascript
var SpreadsheetHandlebars = require('@momsfriendlydevco/spreadsheet-handlebars');

new SpreadsheetHandlebars('input.xlsx')
	.data({...})
	.apply()
	.write('output.xlsx')
```


API
===
The module exposes a single object.


Constructor(options | filename)
-------------------------------
Setup the object either with an options object or a template filename to use.


set(key, [val])
---------------
Set a single or multiple options (if key is an object).
Lodash array and dotted notation is supported for the key.


readTemplate()
--------------
Parse the input template file.
This function is automatically called if constructor is given a filename when initialized.


apply([data])
-------------
Apply the given data (or the data specified in `options.data`) to the loaded template.


json()
------
Convenience function to return the workbook as a JSON object
This will return an object with each key as the sheet ID and a 2D array of cells
