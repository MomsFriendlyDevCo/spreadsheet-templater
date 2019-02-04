var _ = require('lodash');
var events = require('events');
var util = require('util');
var xlsx = require('xlsx');

function SpreadsheetTemplater(options) {
	// Options {{{
	this.settings = {
		re: {
			expression: /{{(.+?)}}/g,
			repeatStart: /{{#?\s*each\s+(.+?)}}/g,
			repeatEnd: /{{\/each.*?}}/,
		},
		repeaterSilentOnError: false,
		template: {
			path: undefined,
		},
		data: {},
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
	* @returns {SpreadsheetTemplater} This chainable object
	*/
	this.read = path => {
		if (path) this.set('template.path', path);
		this.workbook = xlsx.readFile(this.settings.template.path);
		return this;
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

		if (!this.workbook) throw 'No workbook loaded, use readTemplate() first';

		var repeaters = []; // Repeater replacements we need to make - must be made in reverse order due to the fact we are splicing into an array

		_.forEach(this.workbook.Sheets, (sheet, sheetKey) => {
			_.forEach(sheet, (cell, cellKey) => {
				if (cellKey.startsWith('!') || cell.ignore) return; // Ignore meta-cells or cells we have processed elsewhere

				// Repeaters {{{
				var repeatMatch;
				if (repeatMatch = this.settings.re.repeatStart.exec(cell.v)) {
					// Read horizontally until we hit the repeaterEnd
					var repeater = {
						sheet,
						dataSource: repeatMatch[1],
						start: xlsx.utils.decode_cell(cellKey),
					};

					repeater.items = _.chain(sheet)
						.pickBy((cell, cellKey) => {
							var cellPos = xlsx.utils.decode_cell(cellKey);
							if (cellPos.r != repeater.start.r) return false; // Not on the same row
							if (cellPos.c < repeater.start.c) return false; // Occurs before the starting cell
							return true;
						})
						.forEach((cell, cellKey) => cell.ignore = true)
						.map((cell, cellKey) => ({
							position: xlsx.utils.decode_cell(cellKey),
							...cell,
						}))
						.sortBy('position.c')
						.reduce((t, cell) => { // Stop reading horizontally when we hit the ending tag
							if (t.ended) {
								// Do nothing
							} else if (this.settings.re.repeatEnd.test(cell.v)) {
								t.cells.push(cell); // Append this last cell to the iterables
								t.ended = true; // ... and also end reading
							} else {
								t.cells.push(cell); // Append this cell to the items we iterate over
							}
							return t;
						}, {ended: false, cells: []})
						.get('cells')
						.value();

					repeaters.unshift(repeater);

					return; // Don't process this cell any further
				}
				// }}}

				// Simple expressions - e.g. `{{foo.bar.baz}}` {{{
				cell.v = cell.v.replace(this.settings.re.expression, (match, expression) => {
					cell.w = undefined;
					return _.get(this.settings.data, expression);
				});
				// }}}
			});
		});

		if (repeaters.length) {
			repeaters.forEach(repeater => {
				var data = _.get(this.settings.data, repeater.dataSource);
				if (!_.isArray(data)) {
					if (this.settings.repeaterSilentOnError) {
						data = [];
					} else {
						throw `Cannot use data source "${repeater.dataSource}" as a repeater as it is not an array`;
					}
				}

				data = data.map(dataItem => repeater.items.map(item =>
					item.v
						.replace(this.settings.re.repeatStart, '')
						.replace(this.settings.re.repeatEnd, '')
						.replace(this.settings.re.expression, (match, expression) => _.get(dataItem, expression))
				));

				xlsx.utils.sheet_add_aoa(repeater.sheet, data, {origin: repeater.start});
			});
		}

		return this;
	};
	// }}}

	// json {{{
	/**
	* Convenience function to return the workbook as a JSON object
	* This will return an object with each key as the sheet ID and a 2D array of cells
	* @returns {Object} The current workbook as a JSON object
	*/
	this.json = ()=> {
		return _.mapValues(this.workbook.Sheets, (sheet, sheetKey) => xlsx.utils.sheet_to_json(sheet, {header: 1}))
	};
	// }}}

	// Outputs: write, buffer {{{
	/**
	* Write the template file back to disk
	* @param {string} outputFile The output filename to use
	* @returns {SpreadsheetTemplater} This chainable object
	*/
	this.write = outputFile => {
		if (!this.workbook) throw 'No workbook loaded, use read() first';
		xlsx.writeFile(this.workbook, outputFile);
		return this;
	};


	/**
	* Convenience function to return an Express compatible buffer
	* @param {string} [bookType='xlsx'] The output format to use see https://docs.sheetjs.com/#supported-output-formats for the full list
	*/
	this.buffer = bookType => xlsx.write(this.workbook, {type: 'buffer', bookType: bookType || 'xlsx'});
	// }}}

	// Constructor {{{
	if (_.isString(options)) {
		this.read(options);
	} else if (_.isObject(options)) {
		this.set(options);
	}
	// }}}

	return this;
}

util.inherits(SpreadsheetTemplater, events.EventEmitter);

module.exports = SpreadsheetTemplater;
