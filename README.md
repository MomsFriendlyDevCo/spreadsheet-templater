Spreadsheet-Templater
=====================
Simple templates markup for spreadsheets (via [XLSX](https://docs.sheetjs.com)).

This plugin allows a spreadsheet to use handlebars-like notation to replace cell contents which enables an input spreadsheet to act as a template for incoming data.


```javascript
var SpreadsheetTemplater = require('@momsfriendlydevco/spreadsheet-templater');

new SpreadsheetTemplater('input.xlsx')
	.data({...})
	.apply()
	.write('output.xlsx')
```

See the test directory for some example spreadsheets.


Limitations
-----------
There are a few restrictions with this module, mainly due to time and technical limitations:

* Nested `{{#each}}` + `{{/each}}` statements are not supported
* Due to no support for dynamic row adding, at the time of writing, in the upstream [xlsx-populate](https://github.com/dtjohnson/xlsx-populate) library this module will overwrite all rows below the `{{each}}` blocks with however many rows of data need placing - the spreadsheet contents *below* the `{{each}}` blocks will not be moved down



Markup
======
This module reads all cells in all sheets and applies simple substitutions based on a Handlebars like template based on an input data set.


Simple substitution
-------------------
Simple substitution is performed by putting a [lodash compatible dotted notation](https://lodash.com/docs#get) path inside double braces.
For example `{{people.0.name}}` - extracts from the data object the key `people`, the first element of the array and the subkey `name`.


Repeaters
---------
Basic support is provided for single level repeaters. Repeaters start (`{{#each ITERABLE}}`) in the first cell and are read horizontally until the end is encountered (`{{/each}}`).

For example assuming the following CSV spreadsheet layout:

```
Name,Email,Phone,Address
{{#each people}}{{name}},{{email}},{{phone}},"{{address.street}}, {{address.city}}, {{address.zipcode}}{{/each}}"
```

The spreadsheet would be populated with all items in the `people` collection until exhausted.

If no specific data path is specified for each (i.e. `{{#each}}`) the main data object is assumed to be an array and it is used instead.


API
===
The module exposes a single object.

This module supports the following options:

| Option                  | Type     | Default                    | Description                                                                     |
|-------------------------|----------|----------------------------|---------------------------------------------------------------------------------|
| `re`                    | Object   |                            | The regular expressions used when detecting markup                              |
| `re.expression`         | RegExp   | `/{{(.+?)}}/`              | RegExp to detect a single expression replacement                                |
| `re.repeatStart`        | RegExp   | `/{{#?\s*each\s+(.+?)}}/g` | How to identify the start of a repeater                                         |
| `re.repeatEnd`          | RegExp   | `/{{\/each.*?}}/`          | How to identify the end of a repeater                                           |
| `repeaterSilentOnError` | Boolean  | `false`                    | Whether the module should throw when a non-array path is provided to a repeater |
| `template`              | Object   |                            | Options to control templates                                                    |
| `template.path`         | String   |                            | The source file to process the template from                                    |
| `data`                  | Object   | `{}`                       | The data object used when marking up the template output                        |
| `defaultValue`          | Any      | `''`                       | The value used when no corresponding simple dotted path can be located          |


Constructor([options])
----------------------
Setup the initial object with options.


set(key, [val])
---------------
Set a single or multiple options (if key is an object).
Lodash array and dotted notation is supported for the key.


read([path])
------------
Parse the input template file.
This function is automatically called if constructor is given a filename when initialized.


apply([data])
-------------
Apply the given data (or the data specified in `options.data`) to the loaded template.


json()
------
Convenience function to return the workbook as a JSON object
This will return an object with each key as the sheet ID and a 2D array of cells
