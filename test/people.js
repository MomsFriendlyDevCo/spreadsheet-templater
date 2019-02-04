var _ = require('lodash');
var expect = require('chai').expect;
var faker = require('faker');
var SpreadsheetTemplater = require('..');

describe('Template a speadsheet using people data', ()=> {

	var data;
	before('create fake data', ()=> {
		data = Array.from(new Array(10), ()=> faker.helpers.userCard());
	});

	it('apply the template for a single user', ()=> {
		var result = new SpreadsheetTemplater(`${__dirname}/data/people.xlsx`)
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

	it.skip('apply the template for multiple users', ()=> { // NOT YET SUPPORTED
		var result = new SpreadsheetTemplater(`${__dirname}/data/people.xlsx`)
			.data({people: data})
			.apply()
			.json()

		expect(result.People).to.be.deep.equal(_.flatten([
			['Name', 'Email', 'Phone', 'Address'],
			data.map(p => [p.name, p.email, p.phone, `${p.address.street}, ${p.address.city}, ${p.address.zipcode}`]),
		]));
	});

});
