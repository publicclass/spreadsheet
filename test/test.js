var Spreadsheet = require("../lib/spreadsheet"), assert = require("assert");

exports["A spreadsheet without a key."] = function() {
	assert.throws(function() {
		var s = new Spreadsheet();
	}, Error)
}

exports["Worksheets from a nonexisting spreadsheet"] = function() {
	var s = new Spreadsheet("xyz");
	s.worksheets(
			function(err, ws) {
				assert.equal(err.message,"The spreadsheet at this URL could not be found. Make sure that you have the right URL and that the owner of the spreadsheet hasn&#39;t deleted it.")
			})
}

exports["Worksheets from an existing spreadsheet"] = function(beforeExit) {
	var s = new Spreadsheet("0AlpSPOB54A4UdFN5VjdNaTR5cnY4bXAyZGVtZTdtUkE");
	s.worksheets(function(err, ws) {
		assert.ok(!err)
		assert.ok(ws)
		loaded = true;
	})
	var loaded = false;
	beforeExit(function() {
		assert.ok(loaded)
	})
}

exports["eachRow"] = function(beforeExit) {
	var s = new Spreadsheet("0AlpSPOB54A4UdFN5VjdNaTR5cnY4bXAyZGVtZTdtUkE");
	s.worksheet(1, function(err, ws) {
		assert.ok(!err);
		ws.eachRow(function(err, row, meta) {
			assert.ok(!err);
			assert.type(row, "object");
			assert.type(meta, "object");
			loaded = true;
		})
	})
	var loaded = false;
	beforeExit(function() {
		assert.ok(loaded)
	})
}

exports["eachCell"] = function(beforeExit) {
	var s = new Spreadsheet("0AlpSPOB54A4UdFN5VjdNaTR5cnY4bXAyZGVtZTdtUkE");
	s.worksheet(1, function(err, ws) {
		assert.ok(!err);
		ws.eachCell(function(err, cell, meta) {
			assert.ok(!err);
			assert.type(cell, "object");
			assert.type(meta, "object");
			loaded = true;
		})
	})
	var loaded = false;
	beforeExit(function() {
		assert.ok(loaded)
	})
}

exports["Spreadsheet from URL"] = function() {
	var s = Spreadsheet
			.fromURL("https://spreadsheets.google.com/pub?key=0AlpSPOB54A4UdFN5VjdNaTR5cnY4bXAyZGVtZTdtUkE")
	assert.ok(s instanceof Spreadsheet)
}

exports["Worksheet from URL"] = function() {
	var w = Spreadsheet
			.fromURL("https://spreadsheets.google.com/feeds/list/0AlpSPOB54A4UdFN5VjdNaTR5cnY4bXAyZGVtZTdtUkE/od6")
	assert.ok(w instanceof Spreadsheet.Worksheet)
}

exports["Worksheet row by ID"] = function(beforeExit) {
	var spreadsheet = new Spreadsheet(
			"0AlpSPOB54A4UdFN5VjdNaTR5cnY4bXAyZGVtZTdtUkE");
	spreadsheet.worksheet("od6", function(err, w) {
		assert.ifError(err);
		assert.ok(w instanceof Spreadsheet.Worksheet)
		w.row("cpzh4", function(err, row, meta) {
			assert.ifError(err);
			assert.ok(row instanceof Spreadsheet.Row)
			assert.ok(meta instanceof Spreadsheet.Meta)
			loaded = true;
		})
	})

	var loaded = false;
	beforeExit(function() {
		assert.ok(loaded)
	})
}

exports["Worksheet cell by ID"] = function(beforeExit) {
	var spreadsheet = new Spreadsheet(
			"0AlpSPOB54A4UdFN5VjdNaTR5cnY4bXAyZGVtZTdtUkE");
	spreadsheet.worksheet("od6", function(err, w) {
		assert.ifError(err);
		assert.ok(w instanceof Spreadsheet.Worksheet)
		w.cell("R1C1", function(err, cell, meta) {
			assert.ifError(err);
			assert.ok(cell instanceof Spreadsheet.Cell)
			assert.ok(meta instanceof Spreadsheet.Meta)
			loaded = true;
		})
	})

	var loaded = false;
	beforeExit(function() {
		assert.ok(loaded)
	})
}

exports["Load Row from URL"] = function(beforeExit) {
	var loaded = false;
	Spreadsheet
			.fromURL(
					"https://spreadsheets.google.com/feeds/list/0AlpSPOB54A4UdFN5VjdNaTR5cnY4bXAyZGVtZTdtUkE/od6/public/values/cpzh4",
					function(err, row, meta) {
						assert.ifError(err);
						assert.ok(row instanceof Spreadsheet.Row)
						assert.ok(meta instanceof Spreadsheet.Meta)
						loaded = true;
					})
	beforeExit(function() {
		assert.ok(loaded);
	})
}

exports["Load Cell from URL"] = function(beforeExit) {
	var loaded = false;
	Spreadsheet
			.fromURL(
					"https://spreadsheets.google.com/feeds/cells/0AlpSPOB54A4UdFN5VjdNaTR5cnY4bXAyZGVtZTdtUkE/od6/public/values/R1C1",
					function(err, cell, meta) {
						assert.ifError(err);
						assert.ok(cell instanceof Spreadsheet.Cell)
						assert.ok(meta instanceof Spreadsheet.Meta)
						loaded = true;
					})
	beforeExit(function() {
		assert.ok(loaded);
	})
}

exports["Worksheet Array"] = function(beforeExit) {
	var s = new Spreadsheet("0AlpSPOB54A4UdFN5VjdNaTR5cnY4bXAyZGVtZTdtUkE");
	s.worksheetArray(function(err, spreadsheet, worksheets) {
		assert.ok(!err);
		assert.ok(spreadsheet);
		assert.ok(worksheets);
		assert.eql(worksheets.length, 4);
		assert.ok(spreadsheet.title);
		assert.eql(spreadsheet.title, 'test_spreadsheet_module');
		assert.ok(spreadsheet.author);
		assert.ok(spreadsheet.updated);

		assert.eql(worksheets[0].title, 'sheet1');
		assert.ok(worksheets[0].updated);

		assert.eql(worksheets[1].title, 'sheet2');
		assert.ok(worksheets[1].updated);

		assert.eql(worksheets[2].title, 'sheet3');
		assert.ok(worksheets[2].updated);

		assert.eql(worksheets[3].title, 'sheet4');
		assert.ok(worksheets[3].updated);

		loaded = true;
	})
	var loaded = false;
	beforeExit(function() {
		assert.ok(loaded)
	})
}

exports["Worksheet row 1"] = function(beforeExit) {
	var s = new Spreadsheet("0AlpSPOB54A4UdFN5VjdNaTR5cnY4bXAyZGVtZTdtUkE");
	s.worksheet(1, function(err, ws) {
		assert.ok(!err);
		assert.ok(ws);
		assert.ok(ws.title);
		ws.eachRow(function(err, row, meta) {
			assert.ok(!err);
			assert.type(row, "object");
			assert.type(meta, "object");
			loaded = true;
			count += 1;
		})
	})

	var loaded = false, count = 0;
	beforeExit(function() {
		assert.ok(loaded);
		assert.eql(count, 30);
	})
}

exports["Worksheet row 2"] = function(beforeExit) {
	var s = new Spreadsheet("0AlpSPOB54A4UdFN5VjdNaTR5cnY4bXAyZGVtZTdtUkE");
	s.worksheet(2, function(err, ws) {
		assert.ok(!err);
		ws.eachRow(function(err, row, meta) {
			assert.ok(!err);
			assert.type(row, "object");
			assert.type(meta, "object");
			assert.ok(row.a);
			assert.ok(row.b);
			assert.ok(!row.c); // c is empty
			assert.ok(row.d);
			loaded = true;
			count += 1;
		})
	})

	var loaded = false, count = 0;
	beforeExit(function() {
		assert.ok(loaded);
		assert.eql(count, 1);
	})
}

exports["Worksheet 2 eachCell"] = function(beforeExit) {
	var s = new Spreadsheet("0AlpSPOB54A4UdFN5VjdNaTR5cnY4bXAyZGVtZTdtUkE");
	s.worksheet(2, function(err, ws) {
		assert.ok(!err);
		ws.eachCell(function(err, cell, meta) {
			assert.ok(!err);
			assert.type(cell, "object");
			assert.type(meta, "object");
			loaded = true;
			count += 1;
		})
	})

	var loaded = false, count = 0;
	beforeExit(function() {
		assert.ok(loaded);
		assert.eql(count, 7);
	})
}


exports["Worksheet row 3"] = function(beforeExit) {
	var s = new Spreadsheet("0AlpSPOB54A4UdFN5VjdNaTR5cnY4bXAyZGVtZTdtUkE");
	s.worksheet(3, function(err, ws) {
		assert.ok(!err);
		ws.eachRow(function(err, row, meta) {
			assert.equal(err.message, "No rows found.");
			loaded = true;
		})
	})

	var loaded = false;
	beforeExit(function() {
		assert.ok(loaded);
	})
}

exports["Worksheet row 4"] = function(beforeExit) {
	var s = new Spreadsheet("0AlpSPOB54A4UdFN5VjdNaTR5cnY4bXAyZGVtZTdtUkE");
	s.worksheet(4, function(err, ws) {
		assert.ok(!err);
		ws.eachRow(function(err, row, meta) {
			assert.equal(err.message, "No rows found.");
			loaded = true;
		})
	})

	var loaded = false;
	beforeExit(function() {
		assert.ok(loaded);
	})
}

exports["Worksheet row with an invalid id"] = function(beforeExit) {
	var s = new Spreadsheet("0AlpSPOB54A4UdFN5VjdNaTR5cnY4bXAyZGVtZTdtUkE");
	s.worksheet('invalid', function(err, ws) {
			assert.equal(err.message, "A Worksheet must have an valid id.");
			loaded = true;
	})

	var loaded = false;
	beforeExit(function() {
		assert.ok(loaded);
	})
}

exports["Worksheet row with an invalid index id"] = function(beforeExit) {
	var s = new Spreadsheet("0AlpSPOB54A4UdFN5VjdNaTR5cnY4bXAyZGVtZTdtUkE");
	s.worksheet(100, function(err, ws) {
			assert.equal(err.message, "A Worksheet must have an valid id.");
			loaded = true;
	})

	var loaded = false;
	beforeExit(function() {
		assert.ok(loaded);
	})
}