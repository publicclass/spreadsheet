/**
*	Google Spreadsheets API for CommonJS
*/

var xml2js = require("xml2js-expat")
	, request = require("request");

/**
*  A Google Spreadsheet.
*  
*  @param {String} key      The spreadsheet key.
*  @constructor
*/
function Spreadsheet(key){
  if( !key )
    throw new Error("A Spreadsheet must have a key.");
  this.key = key;
}

/**
*  Retrieve all worksheets of the Spreadsheet.
*  
*  @param {Function} fn     Function called for each Worksheet: `function(err,worksheet){}`
*  @return Itself for chaining.
*/
Spreadsheet.prototype.worksheets = function(fn){
  var sheet = this;
  return this.load(function(err,atom){
    if( err ) return fn(err);
    // Create a worksheet from each entry
    var entries = Array.isArray(atom.entry) ? atom.entry : [atom.entry];
    entries.forEach(function(entry,i){
      fn(null,new Worksheet(sheet,entry.id.match(/\w+$/)[0],i+1,entry.title["#"]))
    })
  });
}


/**
*  Retrieve a worksheet by worksheetID.
*  
*  @param {String} id       The ID of the worksheet.
*  @param {Function} fn     Function called with a worksheet instance.
*  @return Itself for chaining.
*/
Spreadsheet.prototype.worksheet = function(id,fn){
  fn(null,new Worksheet(this,id))
  return this;
}

/**
*  Loads the Spreadsheet from Google Data API.
*  
*  @param {Function} fn     Called when loaded.
*  @return Itself for chaining.
*  @private
*/
Spreadsheet.prototype.load = function(fn){ // type should be "list" or "cells"
  atom(["http://spreadsheets.google.com/feeds/worksheets",this.key,"public/values"].join("/"),fn);
  return this;
}

/**
*  Convenience method for creating a Spreadsheet from a shared url.
*  
*  @param {String} url      The URL from which the key is extracted.
*  @return A Spreadsheet instance. Or null if a Spreadsheet could not be created from that URL.
*/
Spreadsheet.fromURL = function(url){
  var key = (require("url").parse(url,true).query||{}).key;
  if( key ) return new Spreadsheet(key);
  return null;
}


/**
*  Each Spreadsheet contains at least one worksheet. 
*  
*  @param {String|Number} id  The worksheetID or page number of the worksheet.
*  @constructor
*/
function Worksheet(spreadsheet,id,index,title){
  if( !spreadsheet )
    throw new Error("A Worksheet must belong to a Spreadsheet.");
  if( !id )
    throw new Error("A Worksheet must have an id.");
  this.id = id;
  this.spreadsheet = spreadsheet;
  this.index = typeof id == "number" ? id : index;
  this.title = title;
}

/**
*  Go through each row in the worksheet.
*  
*  @param {Function} fn     Function called for each row: function(err,row,meta){}
*  @return Itself for chaining.
*/
Worksheet.prototype.eachRow = function(fn){
  return this.load("list",function(err,atom){
    if( err ) return fn(err);
    if( !atom.entry ) return fn(new Error("No rows found."));
    var entries = Array.isArray(atom.entry) ? atom.entry : [atom.entry];
    entries.forEach(function(entry,i){
      // Build a row object from each entrys gdx: items.
      var row = {}
      for( var k in entry )
        if( k.indexOf("gsx:") == 0 )
          row[k.slice(4)] = entry[k];
          
      // And a meta object from the entrys 'id','updated', 'index' and 'total'
      var meta = {
        id: entry.id,
        updated: new Date(entry.updated),
        index: parseInt(atom["openSearch:startIndex"])+i,
        total: parseInt(atom["openSearch:totalResults"])
      }
      fn(null,row,meta);
    })
  });
}

/**
*  Go through each cell in the worksheet.
*  
*  @param {Function} fn     Function called for each cell: function(err,cell,meta){}
*  @return Itself for chaining.
*/
Worksheet.prototype.eachCell = function(fn){
  var ws = this;
  return this.load("cells",function(err,atom){
    if( err ) return fn(err);
    var entries = Array.isArray(atom.entry) ? atom.entry : [atom.entry];
    entries.forEach(function(entry,i){
      // Build a cell object from each entrys gdx: items. 
      var cell = {
        row: parseInt(entry['gs:cell']['@'].row),
        col: parseInt(entry['gs:cell']['@'].col),
        value: entry['gs:cell']['#']
      }
      // And a meta object from the entrys 'id','updated', 'index' and 'total'
      var meta = {
        id: entry.id,
        updated: new Date(entry.updated),
        index: parseInt(atom["openSearch:startIndex"])+i,
        total: parseInt(atom["openSearch:totalResults"])
      }
      fn(null,cell,meta);
    })
  });
}

/**
*  Loads the Worksheet from Google Data API.
*  
*  @param {String} type     Type of feed. Should be "list" or "cells".
*  @param {Function} fn     Called when loaded.
*  @return Itself for chaining.
*  @private
*/
Worksheet.prototype.load = function(type,fn){
  atom(["http://spreadsheets.google.com/feeds",type,this.spreadsheet.key,this.id,"public/values"].join("/"),fn);
  return this;
}

/**
*  Helper method for retrieving entries from an atom feed.
*  
*  @param {String} url      The URL to the atom feed.
*  @param {Function} fn     A callback like: function(err,atom){}
*  @private
*/
function atom(url,fn){
  request({uri:url}, function(err,res,body) {
    if( err ) return fn(err);
		if( res.statusCode == 200 ) {
			var parser = new xml2js.Parser();
			parser.on("end",function(root){
			  fn(null,root);
			}).on("error",function(err){
			  fn(err);
			}).parse(body);
		} else {
		  fn(new Error(body))
		}
	})
}

module.exports = Spreadsheet;