var _ = require('lodash');
var xlsx = require('xlsx-populate');

function SpreadsheetTemplater(options) {
	// Options {{{
	this.settings = {
		re: {
			expression: /{{(.+?)}}/g,
			repeatStart: /{{#?\s*each\s*(.*?)}}/g,
			repeatEnd: /{{\/each.*?}}/,
		},
		repeaterSilentOnError: false,
		template: {
			path: undefined,
		},
		data: {},
		defaultValue: '',
	};

	/**
	* Set a single, or multiple options
	* @param {Object|string} key Either an options object to merge or a single key path (dotted / array notation supported) to set
	* @param {*} [val] If key is a string this specifies the value to set
	* @returns {SpreadsheetHTemplater} This chainable object
	*/
	this.set = function(key, val) {
		if (_.isObject(key)) {
			_.merge(this.settings, key);
		} else {
			_.set(this.settings, key, val);
		}
		return this;
	};

	/**
	* Convenience function to set data
	* @param {Object} data The data to set
	* @returns {SpreadsheetTemplater} This chainable object
	*/
	this.data = data => this.set('data', data);
	// }}}

	// read {{{
	/**
	* The in-memory XLSX workbook
	* @var {xlsx.workbook}
	*/
	this.workbook;


	/**
	* Read the template file specified in settings.templatePath into memory
	* @param {string} [path] Optional path to read, if specified settings.template.path is set, if unspecified the former is used
	* @returns {Promise} Will resolve with this object when completed
	*/
	this.read = path => {
		if (path) this.set('template.path', path);
		return xlsx.fromFileAsync(this.settings.template.path)
			.then(workbook => this.workbook = workbook)
			.then(()=> this)
	};
	// }}}

	// apply {{{
	/**
	* Apply the given data to the template
	* @param {Object} [data] Optional data to set (overriding options.data)
	* @returns {SpreadsheetTemplater} This chainable object
	*/
	this.apply = data => {
		if (data) this.set('data', data);

		if (!this.workbook) throw 'No workbook loaded, use read() first';

		var repeaters = []; // Repeater replacements we need to make - must be made in reverse order due to the fact we are splicing into an array

		this.workbook.sheets().forEach(sheet => {
			var range = sheet.usedRange();
			var endCell = range.endCell();
			range.forEach(cell => {
				if (cell.ssTemplaterIgnore) return; // Cell marked as ignored

				// Repeaters {{{
				var repeatMatch;
				if (repeatMatch = this.settings.re.repeatStart.exec(cell.value())) {
					// Read horizontally until we hit the repeaterEnd
					var repeater = {
						sheet,
						dataSource: repeatMatch[1],
						col: cell.columnNumber(),
						row: cell.rowNumber(),
						range: sheet.range(cell.rowNumber(), cell.columnNumber(), cell.rowNumber(), endCell.columnNumber()),
					};

					// Move ending of range inwards to closing repeater marker
					var closingCell;
					repeater.range.forEach(cell => {
						if (!closingCell && this.settings.re.repeatEnd.test(cell.value())) closingCell = cell;
					});
					if (closingCell) repeater.range = sheet.range(cell.rowNumber(), cell.columnNumber(), cell.rowNumber(), closingCell.columnNumber());

					// Mark cells as ignored so the simple expression replacement doesn't fire
					range.forEach(cell => cell.ssTemplaterIgnore = true);

					// Add it to the start of the repeater array (we need to action these upside down when appending to an array)
					repeaters.unshift(repeater);

					return; // Don't process this cell any further
				}
				// }}}

				// Simple expressions - e.g. `{{foo.bar.baz}}` {{{
				cell.find(this.settings.re.expression, (match, expression) => {
					cell.w = undefined;
					return _.get(this.settings.data, expression, this.settings.defaultValue);
				});
				// }}}
			});
		});

		if (repeaters.length) {
			repeaters.forEach(repeater => {
				var data =
					repeater.dataSource ? _.get(this.settings.data, repeater.dataSource, this.settings.defaultValue) // Use a dotted path as the source
					: this.settings.data // Probably in format `{{#each}}` (use global object)

				if (!_.isArray(data)) {
					if (this.settings.repeaterSilentOnError) {
						data = [];
					} else {
						throw `Cannot use data source "${repeater.dataSource || 'ROOT'}" as a repeater as it is not an array`;
					}
				}

				// Calculate the contents of the range we are replacing
				repeater.rangeTemplate = repeater.range.map(cell => cell.value())[0];

				// Grow repeater.range to have enough space for all the data rows
				repeater.range = repeater.sheet.range(repeater.range.startCell().rowNumber(), repeater.range.startCell().columnNumber(), repeater.range.endCell().rowNumber() + data.length - 1, repeater.range.endCell().columnNumber());
				repeater.range.value(cell => {
					var rowOffset = cell.rowNumber() - repeater.range.startCell().rowNumber();
					var colOffset = cell.columnNumber() - repeater.range.startCell().columnNumber();

					return repeater.rangeTemplate[colOffset]
						.replace(this.settings.re.repeatStart, '')
						.replace(this.settings.re.repeatEnd, '')
						.replace(this.settings.re.expression, (match, expression) => _.get(data, rowOffset + '.' + expression, this.settings.defaultValue))
				});
			});
		}

		return this;
	};
	// }}}

	// json {{{
	/**
	* Convenience function to return the workbook as a JSON object
	* This will return an object with each key as the sheet ID and a 2D array of cells
	* NOTE: This function automatically prunes undefined values from the end of row and cell set
	* @returns {Object} The current workbook as a JSON object
	*/
	this.json = ()=> {
		return _(this.workbook.sheets())
			.mapKeys(sheet => sheet.name())
			.mapValues(sheet =>
				sheet.usedRange().value().map(row =>
					_.dropRightWhile(row, y => !y)
				)
			)
			.value()
	};
	// }}}

	// Outputs: write, buffer {{{
	/**
	* Write the template file back to disk
	* @param {string} outputFile The output filename to use
	* @returns {Promise} A promise which will resolve with this instance
	*/
	this.write = outputFile => {
		if (!this.workbook) throw 'No workbook loaded, use read() first';
		return this.workbook.toFileAsync(outputFile)
			.then(()=> this)
	};


	/**
	* Convenience function to return an Express compatible buffer
	* @returns {Promise} A promise which will resolve with the output buffer contents
	*/
	this.buffer = ()=> {
		if (!this.workbook) throw 'No workbook loaded, use read() first';
		return this.workbook.outputAsync('buffer');
	};
	// }}}

	// Constructor {{{
	if (_.isString(options)) {
		this.set('template.path', options);
	} else if (_.isObject(options)) {
		this.set(options);
	}
	// }}}

	return this;
}

module.exports = SpreadsheetTemplater;
