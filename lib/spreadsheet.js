/**
*  Google Spreadsheets API for CommonJS
*/

var xml2js = require("xml2js-expat")
  , open = require("open-uri");

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
Spreadsheet.prototype.load = function(fn){
  atom(["http://spreadsheets.google.com/feeds/worksheets",this.key,"public/values"].join("/"),fn);
  return this;
}

/**
*  Convenience method for creating a Spreadsheet from a shared url.
*  
*  @param {String} url      The URL from which the key is extracted.
*  @param {Function} fn     The callback for when the URL is a cell- or row-id.
*  @return A Spreadsheet or Worksheet instance. Or null if neither could not be created from that URL.
*/
Spreadsheet.fromURL = function(url,fn){
  url = require("url").parse(url,true);
  var key = (url.query||{}).key;
  if( key ) return new Spreadsheet(key);
  if( md = RE_FEED_URL.exec(url.pathname) ){
    var s = new Spreadsheet(md[2]);
    if( !md[3] ) return s;
    var w = new Worksheet(s,md[3]);
    if( !md[4] ) return w;
    if( fn && md[1] == "list" )
      return w.row(md[4],fn);      
    if( fn && md[1] == "cells" )
      return w.cell(md[4],fn);
  }
  return null;
}

var RE_FEED_URL = /feeds\/(list|cells)\/([\w\d]+)(?:\/([\w\d]+))?(?:\/public\/values\/([\w\d]+))?/i;

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
      fn(null,Row(entry),Meta(entry,atom,i));
    })
  })
}

/**
*  Finds a single Row in a Worksheet by id.
*  
*  @param {String} id       The ID of the single row (returned in the meta of an eachRow callback)
*  @param {Function} fn     Function called for each row: function(err,row,meta){}
*  @return Itself for chaining.
*/
Worksheet.prototype.row = function(id,fn){
  return this.load("list",id,function(err,entry){
    if( err ) return fn(err);
    if( !entry ) return fn(new Error("No row found with id:"+id));
    fn(null,Row(entry),Meta(entry));
  })
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
      fn(null, Cell(entry),Meta(entry,atom,i));
    })
  });
}

/**
*  Finds a single Cell in a Worksheet by id.
*  
*  @param {String} id       The ID of the single cell (returned in the meta of an eachRell callback)
*  @param {Function} fn     Function called for each cell: function(err,cell,meta){}
*  @return Itself for chaining.
*/
Worksheet.prototype.cell = function(id,fn){
  return this.load("cells",id,function(err,entry){
    if( err ) return fn(err);
    if( !entry ) return fn(new Error("No cell found with id:"+id));
    fn(null,Cell(entry),Meta(entry));
  })
}

/**
*  Loads the Worksheet from Google Data API.
*  
*  @param {String} type     Type of feed. Should be "list" or "cells".
*  @param {Function} fn     Called when loaded.
*  @return Itself for chaining.
*  @private
*/
Worksheet.prototype.load = function(type,id,fn){
  if( typeof id == "function" ) fn = id, id = "";
  else id = "/"+id;
  atom(["http://spreadsheets.google.com/feeds",type,this.spreadsheet.key,this.id,"public/values"].join("/")+id,fn);
  return this;
}

/**
*  Instance of a Row
*  
*  @param {Object} entry    An Atom Entry to extract the rows fields from
*/
function Row(entry){
  if( !(this instanceof Row) )
    return new Row(entry);
  // Build a row object from each entrys gdx: items.
  for( var k in entry )
    if( k.indexOf("gsx:") == 0 )
      this[k.slice(4)] = entry[k];
}


/**
*  Instance of a Cell
*  
*  @param {Object} entry    An Atom Entry to extract the cell contents from
*/
function Cell(entry){
  if( !(this instanceof Cell) )
    return new Cell(entry);
  this.row = parseInt(entry['gs:cell']['@'].row)
  this.col = parseInt(entry['gs:cell']['@'].col)
  this.value = entry['gs:cell']['#']
}

/**
*  Instance of a Meta
*  @private
*/
function Meta(entry,atom,i){
  if( !(this instanceof Meta) )
    return new Meta(entry,atom,i);
  // And a meta object from the entrys 'id','updated', 'index' and 'total'
  this.id = entry.id;
  this.updated = new Date(entry.updated);
  if( atom ){
    this.index = parseInt(atom["openSearch:startIndex"])+i;
    this.total = parseInt(atom["openSearch:totalResults"]);
  }
}

/**
*  Helper method for retrieving entries from an atom feed.
*  
*  @param {String} url      The URL to the atom feed.
*  @param {Function} fn     A callback like: function(err,atom){}
*  @private
*/
function atom(url,fn){
  open(url+"?hl=en",function(err,body,res) {
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

// Expose the other constructors for testing
Spreadsheet.Worksheet = Worksheet;
Spreadsheet.Row = Row;
Spreadsheet.Cell = Cell;
Spreadsheet.Meta = Meta;