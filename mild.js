var fs = require("fs");
var JSZip = require("jszip");
var xmldom = require("xmldom");
var parser = new xmldom.DOMParser();
var serializer = new xmldom.XMLSerializer();
var _ = require("lodash");

(function(){
	// adding convinience methods to xmldom inner classes
	var doc = (new xmldom.DOMImplementation()).createDocument();
	var elem = doc.createElement("doesntreallymatter");
	elem.constructor.prototype.attr = function(name, value){
		if(_.isUndefined(value)){
			return this.getAttribute(name);
		}else{
			this.setAttribute(name, value);
		}
	}
	doc.constructor.prototype.byTag = elem.constructor.prototype.byTag = elem.constructor.prototype.getElementsByTagName;
	doc.constructor.prototype.onlyTag = elem.constructor.prototype.onlyTag = function(tagName){
		var list = this.getElementsByTagName(tagName);
		if(list.length <= 1){
			return list[0];
		}else{
			throw new Error("There are " + list.length + " elements with tag name <" + tagName + ">, only one expected");
		}
	}
	elem.constructor.prototype.delete = function(){
		this.parentNode.removeChild(this);
	}
	var nodeList = doc.childNodes;
	nodeList.constructor.prototype.find = function(query){
		return _.find(this, function(node){
			return _.every(query, function(val, key){
				return node.getAttribute(key) == val;
			})
		})
	}

})();

function XLSXWrapper(filename){
	var xml_regex = /\.xml$/;
	var xml_rels_regex = /\.xml\.rels$/;
	var rels_folder_regex = /\/_rels$/

	function jsformat(name, folderPath, isFolder){
		var name = name.replace(/^\$/, "$$$");
		if(isFolder){
			return name;
		}
		if(rels_folder_regex.test(folderPath) && xml_rels_regex.test(name)){
			return name.replace(xml_rels_regex, "");
		}else if(!rels_folder_regex.test(folderPath) && xml_regex.test(name)){
			return name.replace(xml_regex, "");
		}else{
			return name;
		}
	}

	function FolderWrapper(props){
		var self = this;
		var parent = props.parent;
		var content = {};
		var funcs = {
			delete: function(obj){
				var key = _.findKey(self, obj);
				delete self[key];
				var name = _.findKey(content, obj);
				delete content[name]
			}
		}
		this.$isFolder = true;
		this.$isFile = false;
		this.$ = {};

		this.$path = function(){
			return props.path;
		}
		this.$relative = function(path){
			var folders = props.path ? props.path.split("/") : [];
			var folders2 = path.split("/");
			while(folders.length && folders2.length && folders[0] == folders2[0]){
				folders.shift();
				folders2.shift();
			}
			return folders.map(function(){return ".."}).concat(folders2).join("/");
		}

		this.$create = function(path, data, defaultData){
			var parts = path.split("/")
			var name = parts.shift();
			var rest = parts.join("/");
			if(name == ".."){
				return props.parent.$create(rest, data);
			}
			var jsname = jsformat(name, props.path, !!rest);
			if(rest){
				if(self[jsname] && self[jsname].$isFile){
					throw new Error("Attempted to create a folder instead of existing file");
				}
				if(!self[jsname]){
					self[jsname] = content[name] = new FolderWrapper(
						_.defaults({
							path: props.path ? props.path + "/" + name : name,
							parent: self,
							funcs: funcs
						}, props)
					);
				}
				self[jsname].$create(rest, data, defaultData);
				return self[jsname].$get(rest);
			}else{
				var Wrapper = (xml_regex.test(name) || xml_rels_regex.test(name)) ? XMLWrapper : FileWrapper;
				return self[jsname] = content[name] = new Wrapper(
					_.defaults({
						path: props.path ? props.path + "/" + name : name,
						parent: self,
						funcs: funcs,
						data: data,
						defaultData: defaultData
					}, props)
				);
			}
		}

		this.$get = function(path){
			var parts = path.split("/")
			var name = parts.shift();
			var rest = parts.join("/");
			if(name == ".."){
				if(!parent){
					return null;
				}else{
					return parent.$get(rest);
				}
			}else if(!content[name]){
				return null;
			}else if(!rest){
				return content[name];
			}else{
				return content[name].$get(rest);
			}
		}

		this.$clone = function(src, dest){
			var source = self.$get(src);
			if(_.isUndefined(dest)){
				dest = self.$relative(source.$parent().$path()) + "/" + source.$parent().$uniqueLike(source.$name());
			}
			console.log("clone " + src + " " + dest);
			return self.$create(dest, source.data());
		}
		this.$unique = function(prefix, suffix){
			var i = 0;
			while(prefix + ++i + suffix in content);
			return prefix + i + suffix;
		}
		this.$uniqueLike = function(name){
			var arr = name.split(/[0-9]+/);
			var prefix = arr.shift();
			var suffix = arr.join("");
			return self.$unique(prefix, suffix)
		}
		this.$save = function(){
			_.forEach(content, function(wrapper){
				wrapper.$save();
			})
		}
	}

	function destroy(obj){
		_.forEach(obj, function(value, key){
			if(typeof value == "function"){
				obj[key] = function(){
					throw new Error("You've tried to call a method of destroyed object. Normally it shouldn't happen.");
				}
			}else{
				delete obj[key];
			}
		})
	}

	function FileWrapper(props){
		var self = this;
		var zip = props.zip;
		var path = props.path;
		var data = props.data;
		var defaultData = props.defaultData;

		this.$isFolder = false;
		this.$isFile = true;

		this.delete = function(){
			zip.remove(path);
			props.funcs.delete(self);
			destroy(self);
		}
		this.data = function(){
			data = data || (zip.file(path) ? zip.file(path).asNodeBuffer() : defaultData);
			return data;
		}
		this.$save = function(){
			data && zip.file(path, data);
		}
		this.$path = function(){
			return props.path;
		}
		this.$parent = function(){
			return props.parent;
		}
		this.$name = function(){
			return props.path.split("/").pop();
		}
	}

	function XMLWrapper(props){
		console.log("new xmlwrapper");
		var self = this;
		var zip = props.zip;
		var path = props.path;
		var name = path.split("/").pop();
		var xml = props.data && parser.parseFromString(props.data);

		this.delete = function(){
			zip.remove(path);
			props.funcs.delete(self);
			xml = null;
			destroy(self);
		}
		this.data = function(){
			return serializer.serializeToString(self.xml);
		}
		this.$save = function(){
			xml && zip.file(path, serializer.serializeToString(xml));
		}
		this.$path = function(){
			return props.path;
		}
		this.$parent = function(){
			return props.parent;
		}
		this.$name = function(){
			return props.path.split("/").pop();
		}
		Object.defineProperty(self, "xml", {
			get: function(){
				if(xml){
					return xml;
				}else{
					console.log(props.defaultData);
					return xml = parser.parseFromString(zip.file(path) ? zip.file(path).asText() : props.defaultData);
				}
			}
		})
		Object.defineProperty(self, "rels", {
			get: function(){
				return props.parent.$get("_rels/" + name + ".rels");
			}
		})
	}


	var zip = new JSZip(fs.readFileSync(filename));

	var root = new FolderWrapper({
		zip: zip,
		path: ""
	});

	_.forEach(zip.files, function(file, filepath){
		root.$create(filepath);
	})

	root.$create("xl/sharedStrings.xml", null, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+
		'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>');
	root.$saveFile = function(newname){
		root.$save();
		fs.writeFile(newname || filename, zip.generate({type:"nodebuffer"}));
	}

	return root;
}

function Workbook(filename){
	var xlsx = new XLSXWrapper(filename);
	var wbRelsNode = xlsx.xl.workbook.rels.xml.onlyTag("Relationships");
	var types = xlsx["[Content_Types]"].xml.onlyTag("Types");

	//creating SST if it isn't created before
	xlsx.$create("xl/sharedStrings.xml", null, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+
	'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>');
	var SSTRel = wbRelsNode.childNodes.find({Target: "sharedStrings.xml"});
	if(!SSTRel){
		SSTRel = xlsx.xl.workbook.rels.xml.createElement("Relationship");
		SSTRel.attr("Target", "sharedStrings.xml");
		SSTRel.attr("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings");
		var i = 0;
		while(wbRelsNode.childNodes.find({Id: "rId"+(++i)}));
		var rId = "rId" + i;
		SSTRel.attr("Id", rId);
		wbRelsNode.appendChild(SSTRel);
	}
	var SSTOverride = types.byTag("Override").find({PartName:"/xl/sharedStrings.xml"});
	if(!SSTOverride){
		SSTOverride = xlsx["[Content_Types]"].xml.createElement("Override");
		SSTOverride.attr("PartName", "/xl/sharedStrings.xml");
		SSTOverride.attr("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
		types.appendChild(SSTOverride);
	}

	this.sheets = function(name){
		var sheetsNode = xlsx.xl.workbook.xml.onlyTag("sheets");
		if(!name){
			return _.map(sheetsNode.childNodes, function(node){
				return node.attr("name");
			});
		}else{
			var sheetNode = sheetsNode.childNodes.find({name: name});
			if(!sheetNode){
				return null;
			}
			var sheetRel = xlsx.xl.workbook.rels.xml
				.byTag("Relationship").find({"Id": sheetNode.attr("r:id")});
			if(!sheetRel){
				throw new Error("Invalid xlsx file");
			}
			var sheetXML = xlsx.xl.$get(sheetRel.attr("Target"));
			if(!sheetXML){
				throw new Error("Invalid xlsx file");
			}
			return new Worksheet(xlsx, sheetXML, sheetRel.attr("Target"));
		}
	}
	this.save = function(fname){
		xlsx.$saveFile(fname);
	}
	this.xlsx = xlsx; //debug only
}

function Worksheet(xlsx, sheet, path){
	var self = this;

	this.name = function(val){
		var relations = xlsx.xl.workbook.rels.xml
			.byTag("Relationship");
		var sheetId = relations.find({Target: path}).attr("Id");
		var wbSheetNode = xlsx.xl.workbook.xml.byTag("sheet").find({"r:id": sheetId});

		if(val){
			wbSheetNode.attr("name", val);	
		}else{
			return wbSheetNode.attr("name");
		}
	}
	this.range = function(){
		return new Range(xlsx, sheet, getBounds.call(null, arguments));
	}
	this.clone = function(name){
		var wbSheetRelList = xlsx.xl.workbook.rels.xml.onlyTag("Relationships");
		var wbSheetRelNode = wbSheetRelList.childNodes.find({Target: path});
		var cloneSheetRelNode = wbSheetRelNode.cloneNode();
		
		var i = 0;
		while(wbSheetRelList.childNodes.find({Id: "rId"+(++i)}));
		var rId = "rId" + i;
		cloneSheetRelNode.attr("Id", rId);
		
		var wbSheetList = xlsx.xl.workbook.xml.onlyTag("sheets");
		var sheetId = 0;
		while(wbSheetList.childNodes.find({sheetId: ++sheetId}));
		var wbSheetNode = wbSheetList.childNodes.find({"r:id": wbSheetRelNode.attr("Id")});
		
		var cloneSheet = xlsx.$clone(sheet.$path());
		cloneSheetRelNode.attr("Target", xlsx.xl.$relative(cloneSheet.$path()));
		wbSheetRelList.appendChild(cloneSheetRelNode);

		var cloneSheetNode = wbSheetNode.cloneNode();
		cloneSheetNode.attr("sheetId", sheetId);
		cloneSheetNode.attr("r:id", rId);
		cloneSheetNode.attr("name", name);
		wbSheetList.appendChild(cloneSheetNode);

		var cloneRels = xlsx.$clone(sheet.rels.$path());
		
		var override = xlsx["[Content_Types]"].xml.createElement("Override");
		override.attr("PartName", "/" + cloneSheet.$path());
		override.attr("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
		xlsx["[Content_Types]"].xml.onlyTag("Types").appendChild(override);

		_.forEach(sheet.rels.xml.byTag("Relationship"), function(rel){
			var clonedFile = xlsx.xl.worksheets.$clone(rel.attr("Target"));
			var newTarget = rel.attr("Target").replace(/\/[^/]+$/, "/") + clonedFile.$name();
			cloneRels.xml.byTag("Relationship").find({Id: rel.attr("Id")}).attr("Target", newTarget);
		});
		return new Worksheet(xlsx, cloneSheet, xlsx.xl.$relative(cloneSheet.$path()));
	}
	this.delete = function(){
		var wbSheetRel = xlsx.xl.workbook.rels.xml
			.byTag("Relationship").find({Target: path});
		var wbSheetsCount = xlsx.xl.workbook.xml.byTag("sheet").length;
		if(wbSheetsCount == 1){
			throw new Error("Attempted to delete last sheet");
		}
		var wbSheetNode = xlsx.xl.workbook.xml.byTag("sheet").find({"r:id": wbSheetRel.attr("Id")});
		_.forEach(sheet.rels.xml.byTag("Relationship"), function(rel){
			xlsx.xl.worksheets.$get(rel.attr("Target")).delete();
		})
		xlsx["[Content_Types]"].xml.byTag("Override").find({PartName: "/" + sheet.$path()}).delete();
		sheet.rels.delete();
		sheet.delete();
		wbSheetNode.delete();
		wbSheetRel.delete();
	}
}

var LETTERS = "@ABCDEFGHIJKLMNOPQRSTUVWXYZ";

function addr(arg1, arg2){
	if(arg2){
		return addr(arg1) + arg2;
	}else{
		var q = arg1;
		var r; 
		var s = "";
		while(q > 0){
			r = q % 26 || 26;
			s = LETTERS[r] + s;
			q = (q - r)/26;
		}
		return s;
	}
}

function index(str){
	var letters = (str.match(/[A-Z]+/, str)||[false])[0];
	var digits = (str.match(/[0-9]+/, str)||[false])[0];
	if(digits){
		return [index(letters), digits-0];
	}else{
		return letters.split("").reduce(function(acc, val){
			return acc*26 + LETTERS.indexOf(val);
		}, 0)
	}
}

function getBounds(){
	if(arguments.length < 4){
		if(arguments[1]){
			if(typeof arguments[1] == "string"){
				return getBounds(index(arguments[0]), index(arguments[1]));
			}else{
				return getBounds(arguments[0][0], arguments[0][1], arguments[1][0], arguments[1][1]);
			}
		}else{
			if(typeof arguments[0] == "string"){
				var args = arguments[0].split(":");
				return getBounds.apply(null, args);
			}else{
				return getBounds.apply(null, arguments[0]);
			}
		}
	}
	return {
		left: arguments[0],
		top: arguments[1],
		right: arguments[2],
		bottom: arguments[3],

	}
}

function Range(xlsx, sheet, bounds){
	var columnIndexes = _.range(bounds.left, bounds.right + 1);
	var rowIndexes = _.range(bounds.top, bounds.bottom + 1);
	var SST = xlsx.xl.sharedStrings.xml.onlyTag("sst");

	this.bounds = function(){
		return bounds;
	}
	function safeDeleteFormula(cell){
		var f = cell.onlyTag("f");
		if(!f){
			return;
		}
		var ref = f.attr("ref");
		var si = f.attr("si");
		var formula = f.textContent;
		f.delete();
		if(!ref){
			return;
		}
		var fbounds = getBounds(ref);
		var rowNodes = sheet.xml.byTag("row");
		var flag; //some ugly imperative iteration? yes, it is
		for(var i = fbounds.top; i <= fbounds.bottom; i++){
			var row = rowNodes.find({r:i});
			if(!row){
				continue;
			}
			var cellNodes = row.byTag("c");
			for(var j = fbounds.left; j <= fbounds.right; j++){
				var cell2 = cellNodes.find({r: addr(j, i)});
				if(!cell2){
					continue;
				}
				var f2 = cell2.onlyTag("f");
				if(!f2){
					continue;
				}
				if(f.attr("t") == "shared" && f2.attr("si") == si){
					f2.attr("ref", ref);
					f2.textContent = formula;
					flag = true;
					break;
				}

			}
			if(flag){
				break;
			}
		}
	}
	function safeDelete(cell){
		if(!cell){
			return;
		}
		safeDeleteFormula(cell);
		cell.delete();
	}
	function readIterator(f){
		var rowNodes = sheet.xml.byTag("row");
		return rowIndexes.map(function(r){
			var row = rowNodes.find({r:r});
			if(!row){
				return columnIndexes.map(function(){return null});
			}else{
				var cellNodes = row.byTag("c");
				return columnIndexes.map(function(c){
					var cell = cellNodes.find({r: addr(c, r)});
					return f({
						cell: cell,
						row: row,
						c: c,
						r: r
					});
				});
			}
		})
	}
	function writeIterator(data, f){
		rowIndexes.forEach(function(r,x){
			var rowNodes = sheet.xml.byTag("row");
			var row = rowNodes.find({r:r});
			if(!row){
				row = sheet.xml.createElement("row");
				row.attr("r", r);
				row.attr("spans", bounds.left + ":" + bounds.right);
				var sheetNode = sheet.xml.onlyTag("sheetData")
				var row2 = _.find(sheetNode.childNodes, function(elem){
					return elem.attr("r") > r;
				})
				sheet.xml.onlyTag("sheetData").insertBefore(row, row2); //how does it work if row2 === null? o_O
			}
			columnIndexes.forEach(function(c, y){
				var val = ((data || [])[x] || [])[y];
				var cell = row.childNodes.find({r: addr(c, r)});
				f({
					cell: cell,
					row: row,
					c: c,
					r: r,
					val: val
				});
			})
		})
	}
	this.style = function(data){
		if(data){
			writeIterator(data, function(params){
				var cell = params.cell;
				var row = params.row;
				var val = params.val;
				var c = params.c;
				var r = params.r;
				if(!cell && val){
					cell = sheet.xml.createElement("c");
					cell.attr("r", addr(c, r));
					var cell2 = _.find(row.childNodes, function(elem){
						return index(elem.attr("r"))[0] > c;
					})
					row.insertBefore(cell, cell2);
				}
				if(_.isNull(val)){
					cell.removeAttribute("s");
				}else if(_.isNumber(val) || _.isString(val)){
					cell.attr("s", val);
				}else if(_.isUndefined(val)){
					//do nothing
				}else{
					throw new Error("Unsupported cell style");
				}
			});
		}else{
			return readIterator(function(params){
				var cell = params.cell;
				var row = params.row;
				if(!cell){
					return null;
				}
				return cell.attr("s");
			});
		}
	}
	this.formula = function(data){
		if(data){
			if(_.isString(data)){
				var flag = true;
				var si = -1;
				var sheetData = sheet.xml.onlyTag("sheetData");
				while(sheetData.byTag("f").find({si:++si}));
				writeIterator(null, function(params){
					var cell = params.cell;
					var row = params.row;
					var c = params.c;
					var r = params.r;
					if(!cell){
						cell = sheet.xml.createElement("c");
						cell.attr("r", addr(c, r));
						var cell2 = _.find(row.childNodes, function(elem){
							return index(elem.attr("r"))[0] > c;
						})
						row.insertBefore(cell, cell2);
						cell.appendChild(sheet.xml.createElement("v"))
					}
					safeDeleteFormula(cell);
					var f = sheet.xml.createElement("f");
					cell.appendChild(f);
					if(flag){
						f.attr("ref", addr(bounds.left, bounds.top) + ":" + addr(bounds.right, bounds.bottom));
						f.textContent = data;
						flag = false;
					}
					f.attr("si", si);
					f.attr("t", "shared");
					var v = cell.onlyTag("v");
					if(v){
						v.delete();
					}
				})
			}else if(_.isArray(data)){
				writeIterator(data, function(params){
					var cell = params.cell;
					var row = params.row;
					var val = params.val;
					var c = params.c;
					var r = params.r;
					if(!cell && val){
						cell = sheet.xml.createElement("c");
						cell.attr("r", addr(c, r));
						var cell2 = _.find(row.childNodes, function(elem){
							return index(elem.attr("r"))[0] > c;
						})
						row.insertBefore(cell, cell2);
					}
					if(_.isString(val)){
						safeDeleteFormula(cell);
						var f = sheet.xml.createElement("f");
						cell.appendChild(f);
						f.textContent = val;
						var v = cell.onlyTag("v");
						if(v){
							v.delete();
						}
					}else if(_.isUndefined(val)){
						//do nothing. three times
					}else if(_.isNull(val)){
						safeDeleteFormula(cell);
					}else{
						throw new Error("Unsupported cell formula");
					}
				})
			}else{
				throw new Error("Unsupported data argument");
			}
		}else{
			throw new Error("Not implemented yet");	
		}
	}
	this.value = function(data){
		if(data){
			writeIterator(data, function(params){
				var cell = params.cell;
				var row = params.row;
				var val = params.val;
				var c = params.c;
				var r = params.r;
				if(!cell && val){
					cell = sheet.xml.createElement("c");
					cell.attr("r", addr(c, r));
					var cell2 = _.find(row.childNodes, function(elem){
						return index(elem.attr("r"))[0] > c;
					})
					row.insertBefore(cell, cell2);
					cell.appendChild(sheet.xml.createElement("v"))
				}
				if(typeof val == "number"){
					cell.removeAttribute("t");
					var v = cell.onlyTag("v");
					if(!v){
						v = sheet.xml.createElement("v");
						cell.appendChild(v);
					}
					v.textContent = val;
				}else if(typeof val == "string" && val){
					cell.attr("t", "s");
					var n = _.findIndex(SST.childNodes, function(si){
						return si.onlyTag("t").textContent == val;
					});
					var v = cell.onlyTag("v");
					if(!v){
						v = sheet.xml.createElement("v");
						cell.appendChild(v);
					}
					if(n > -1){
						v.textContent = n;
					}else{
						v.textContent = SST.childNodes.length;
						var si = sheet.xml.createElement("si");
						var t = sheet.xml.createElement("t");
						t.textContent = val;
						SST.appendChild(si);
						si.appendChild(t);
					}
				}else if(_.isNull(val)){
					safeDelete(cell);
				}else if(val === ""){
					if(cell){
						var v = cell.onlyTag("v");
						if(v){
							cell.removeChild(v);
						}
					}
				}else if(_.isUndefined(val)){
					//do nothing. twice
				}else{
					throw new Error("Unsupported cell value");
				}
			})
		}else{
			return readIterator(function(params){
				var cell = params.cell;
				var row = params.row;
				if(!cell){
					return null;
				}
				var v = cell.onlyTag("v");
				if(!v){
					return "";
				}
				var val = v.textContent;
				if(cell.attr("t") == "s"){
					val = SST.childNodes[val].onlyTag("t").textContent;
				}
				return val;
			});
		}
	}
}

module.exports = {
	Workbook: Workbook,
	xlsx: XLSXWrapper
}