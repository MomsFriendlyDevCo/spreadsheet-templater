var _ = require('lodash');
var events = require('events');
var util = require('util');
var xlsx = require('xlsx');

function SpreadsheetHandlebars(options) {
	// Options {{{
	this.settings = {
		re: {
			expression: /{{(.+?)}}/g,
		},
		template: {
			path: undefined,
		},
		data: {},
	};

	/**
	* Set a single, or multiple options
	* @param {Object|string} key Either an options object to merge or a single key path (dotted / array notation supported) to set
	* @param {*} [val] If key is a string this specifies the value to set
	* @returns {SpreadsheetHandlebars} This chainable object
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
	* @returns {SpreadsheetHandlebars} This chainable object
	*/
	this.data = data => this.set('data', data);
	// }}}

	// readTemplate {{{
	/**
	* The in-memory XLSX workbook
	* @var {xlsx.workbook}
	*/
	this.workbook;


	/**
	* Read the template file specified in settings.templatePath into memory
	* @returns {SpreadsheetHandlebars} This chainable object
	*/
	this.readTemplate = ()=> {
		this.workbook = xlsx.readFile(this.settings.template.path);
		return this;
	};
	// }}}

	// applyTemplate {{{
	/**
	* Apply the given data to the template
	* @param {Object} [data] Optional data to set (overriding options.data)
	* @returns {SpreadsheetHandlebars} This chainable object
	*/
	this.apply = data => {
		if (data) this.set('data', data);

		if (!this.workbook) throw 'No workbook loaded, use readTemplate() first';
		_.forEach(this.workbook.Sheets, (sheet, sheetKey) => {
			_.forEach(sheet, (cell, cellKey) => {
				if (cellKey.startsWith('!')) return; // Ignore meta-cells

				// Simple expressions - e.g. `{{foo.bar.baz}}`
				cell.v = cell.v.replace(this.settings.re.expression, (match, expression) => {
					cell.w = undefined;
					return _.get(this.settings.data, expression);
				});
			});
		});

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

	// write {{{
	/**
	* Write the template file back to disk
	* @param {string} outputFile The output filename to use
	* @returns {SpreadsheetHandlebars} This chainable object
	*/
	this.write = outputFile => {
		if (!this.workbook) throw 'No workbook loaded, use readTemplate() first';
		xlsx.writeFile(this.workbook, outputFile);
		return this;
	};
	// }}}

	// Constructor {{{
	if (_.isString(options)) {
		this.set('template.path', options);
		this.readTemplate();
	} else if (_.isObject(options)) {
		this.set(options);
	}
	// }}}

	return this;
}

util.inherits(SpreadsheetHandlebars, events.EventEmitter);

module.exports = SpreadsheetHandlebars;
