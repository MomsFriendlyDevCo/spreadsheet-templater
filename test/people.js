var expect = require('chai').expect;
var faker = require('faker');
var SpreadsheetHandlebars = require('..');

describe('Template a speadsheet using people data', ()=> {

	var data;
	before('create fake data', ()=> {
		data = Array.from(new Array(10), ()=> faker.helpers.userCard());
	});

	it('apply the template using the data set', ()=> {
		var result = new SpreadsheetHandlebars(`${__dirname}/data/people.xlsx`)
			.data({people: data})
			.apply()
			.json()

		expect(result.Person).to.be.deep.equal([
			['Example Person'],
			['Name', data[0].name],
			['Email', data[0].email],
			['Address', `${data[0].address.street}, ${data[0].address.city}, ${data[0].address.zipcode}`],
			['Phone', data[0].phone],
		]);
	});

});
