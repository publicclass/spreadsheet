var Spreadsheet = require("../lib/spreadsheet")
  , assert = require("assert");

exports["A spreadsheet without a key."] = function(){
  assert.throws(function(){
    var s = new Spreadsheet();
  },Error)
}

exports["Worksheets from a nonexisting spreadsheet"] = function(){
  var s = new Spreadsheet("xyz");
  s.worksheets(function(err,ws){
    assert.equal(err.message,"The spreadsheet at this URL could not be found. Make sure that you have the right URL and that the owner of the spreadsheet hasn&#39;t deleted it.")
  })
}

exports["Worksheets from an existing spreadsheet"] = function(beforeExit){
  var s = new Spreadsheet("0AkUwNDE9ftXvdGMxQ2hMS2oySG9HeV9OeUtkZ2xoM1E");
  s.worksheets(function(err,ws){
    assert.ok(!err)
    assert.ok(ws)
    loaded = true;
  })
  var loaded = false;
  beforeExit(function(){
    assert.ok(loaded)
  })
}

exports["eachRow"] = function(beforeExit){
  var s = new Spreadsheet("0AkUwNDE9ftXvdGMxQ2hMS2oySG9HeV9OeUtkZ2xoM1E");
  s.worksheet(1,function(err,ws){
    assert.ok(!err);
    ws.eachRow(function(err,row,meta){
      assert.ok(!err);
      assert.type(row,"object");
      assert.type(meta,"object");
      loaded = true;
    })
  })
  var loaded = false;
  beforeExit(function(){
    assert.ok(loaded)
  })
}

exports["eachCell"] = function(beforeExit){
  var s = new Spreadsheet("0AkUwNDE9ftXvdGMxQ2hMS2oySG9HeV9OeUtkZ2xoM1E");
  s.worksheet(1,function(err,ws){
    assert.ok(!err);
    ws.eachCell(function(err,cell,meta){
      assert.ok(!err);
      assert.type(cell,"object");
      assert.type(meta,"object");
      loaded = true;
    })
  })
  var loaded = false;
  beforeExit(function(){
    assert.ok(loaded)
  })
}

exports["Spreadsheet from URL"] = function(){
  var s = Spreadsheet.fromURL("https://spreadsheets.google.com/pub?key=0AkUwNDE9ftXvdGMxQ2hMS2oySG9HeV9OeUtkZ2xoM1E")
  assert.notEqual(s,null,"Spreadsheet should not be null.")
}