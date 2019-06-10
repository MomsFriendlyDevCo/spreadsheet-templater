var _ = require('lodash');
var expect = require('chai').expect;
var faker = require('faker');
var mlog = require('mocha-logger');
var moment = require('moment');
var SpreadsheetTemplater = require('..');
var temp = require('temp');

describe('Template a speadsheet using people data', ()=> {

	var data;
	before('create fake data', ()=> {
		data = Array.from(new Array(10), ()=> ({
			...faker.helpers.userCard(),
			dob: moment()
				.subtract(_.random(0, 50), 'y')
				.subtract(_.random(0, 365), 'd')
				.startOf('day')
				.toDate(),
		}));
	});

	it('should read and output a file without changes', async ()=> {
		var template = await new SpreadsheetTemplater().read(`${__dirname}/data/people.xlsx`);
		var result = template.json();

		expect(result.Person).to.be.deep.equal([
			['Example Person'],
			['Name', '{{people.0.name}}'],
			['Email', '{{people.0.email}}'],
			['Address', '{{people.0.address.street}}, {{people.0.address.city}}, {{people.0.address.zipcode}}'],
			['Phone', '{{people.0.phone}}'],
		]);

		expect(result.People).to.be.deep.equal([
			['Name', 'Email', 'Phone', 'Address', 'Dob', 'Age'],
			['{{#each people}}{{name}}', '{{email}}', '{{phone}}', '{{address.street}}, {{address.city}}, {{address.zipcode}}', '${dob}', '${Math.floor((new Date().getTime() - dob.getTime()) / (1000 * 60 * 60 * 24 * 365.2425))}{{/each}}'],
		]);
	});

	it('apply the template for a single user', async ()=> {
		var template = await new SpreadsheetTemplater().read(`${__dirname}/data/people.xlsx`);
		var result = template
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

	it('apply the template for multiple users', async ()=> {
		var template = await new SpreadsheetTemplater().read(`${__dirname}/data/people.xlsx`);

		var result = template
			.data({people: data})
			.apply()
			.json()

		var expected = [
			['Name', 'Email', 'Phone', 'Address', 'Dob', 'Age'],
			...data.map(p => [
				p.name,
				p.email,
				p.phone,
				`${p.address.street}, ${p.address.city}, ${p.address.zipcode}`,
				p.dob,
				Math.floor((new Date().getTime() - p.dob.getTime()) / (1000 * 60 * 60 * 24 * 365.2425)),
			]),
		];

		expect(result.People).to.be.deep.equal(expected);
	});

	it('dump the templated output to disk', async ()=> {
		var outputPath = temp.path({suffix: '.xlsx'});
		var template = await new SpreadsheetTemplater().read(`${__dirname}/data/people.xlsx`);

		await template
			.data({people: data})
			.apply()
			.write(outputPath);

		mlog.log('saved file to', outputPath);
	});

});
