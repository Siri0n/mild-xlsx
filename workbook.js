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
	var nodeList = doc.childNodes;
	nodeList.constructor.prototype.find = function(query){
		return _.find(this, function(node){
			return _.every(query, function(val, key){
				return node.getAttribute(key) == val;
			})
		})
	}

})();

var LETTERS = "@ABCDEFGHIJKLMNOPQRSTUVWXYZ";

function set(obj, path, val){
	var keys = path.split("/");
	var ref = obj;
	var i = 0;
	while(i < keys.length - 1){
		if(typeof ref[keys[i]] != "object" || ref[keys[i]] === null){
			ref[keys[i]] = {}
		}
		ref = ref[keys[i++]];
	}
	ref[keys[i]] = val;
}

function get(obj, path){
	var keys = path.split("/");
	var ref = obj;
	var i = 0;
	while(i < keys.length){
		if(typeof ref[keys[i]] != "object" || ref[keys[i]] === null){
			return ref[keys[i]];
		}
		ref = ref[keys[i++]];
	}
	return ref;
}

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
		tl:{
			c: arguments[0],
			r: arguments[1]
		},
		br:{
			c: arguments[2],
			r: arguments[3]
		}
	}
}

function createAccessor(zip, filepath){
	var xml = null;
	return function(command){
		if(!command){
			if(!xml){
				xml = parser.parseFromString(
					zip.file(filepath).asText()
				);
			}
			return xml;
		}else if(command == "save"){
			if(xml){
				zip.file(filepath, serializer.serializeToString(xml));
			}
		}
	}
}

function loadXLSX(filename){
	var zip = new JSZip(fs.readFileSync(filename));
	var result = {};
	_.forEach(zip.files, function(file, filepath){
		if(/\.(xml|rels)$/.test(filepath)){
			set(result, filepath, createAccessor(zip, filepath));
		}
	})
	return function(command){
		if(!command){
			return result;
		}else if(command == "save"){
			fs.writeFile(filename, zip.generate({type:"nodebuffer"}))
		}
	}
}

function Workbook(filename){
	var xlsx = loadXLSX(filename);

	var worksheets = _.map(
		xlsx().xl.worksheets,
		function(accessor, path){
			if(typeof accessor == "function"){
				return new Worksheet(xlsx, accessor, path);
			}else{
				return false;
			}
		}
	);
	worksheets = _.compact(worksheets);
	this.sheets = function(name){
		if(!name){
			return worksheets;
		}else{
			return _.find(worksheets, function(sheet){
				return sheet.name() == name;
			});
		}
	}
	this.save = function(){
		function save(arg){
			if(typeof arg == "function"){
				arg("save");
			}else{
				_.forEach(arg, save);
			}
		}
		save(xlsx());
		xlsx("save");
	}
}

function Worksheet(xlsx, accessor, path){
	this.name = function(val){
		var relations = xlsx().xl._rels["workbook.xml.rels"]()
			.byTag("Relationship");
		var sheetId = relations.find({Target: "worksheets/" + path}).attr("Id");

		var wbSheetNode = xlsx().xl["workbook.xml"]().byTag("sheet").find({"r:id": sheetId});

		if(val){
			wbSheetNode.attr("name", val);	
		}else{
			return wbSheetNode.attr("name");
		}
	}
	this.range = function(){
		return new Range(xlsx, accessor, getBounds.call(null, arguments));
	}
}

function Range(xlsx, accessor, bounds){
	var columns = _.range(bounds.tl.c, bounds.br.c + 1);
	var rows = _.range(bounds.tl.r, bounds.br.r + 1);
	var SST = xlsx().xl["sharedStrings.xml"] ? xlsx().xl["sharedStrings.xml"]().onlyTag("sst") : null;

	this.bounds = function(){
		return bounds;
	}
	this.value = function(data){
		var rowNodes = accessor().byTag("row")
		if(data){
			rows.forEach(function(r, x){
				var rowNode = rowNodes.find({r:r});
				if(!rowNode){
					rowNode = accessor().createElement("row");
					rowNode.attr("r", r);
					//rowNode.attr("spans", bounds.tl.c + ":" + bounds.br.c);
					var sheetNode = accessor().onlyTag("sheetData")
					var rowNode2 = _.find(sheetNode.childNodes, function(elem){
						return elem.attr("r") > r;
					})
					accessor().onlyTag("sheetData").insertBefore(rowNode, rowNode2);
				}
				columns.forEach(function(c, y){
					var val = (data[x] || [])[y];
					var node = rowNode.childNodes.find({r: addr(c, r)});
					if(!node && val){
						node = accessor().createElement("c");
						node.attr("r", addr(c, r));
						var node2 = _.find(rowNode.childNodes, function(elem){
							return index(elem.attr("r"))[0] > c;
						})
						rowNode.insertBefore(node, node2);
						node.appendChild(accessor().createElement("v"))
					}
					if(typeof val == "number"){
						node.removeAttribute("t");
						var v = node.onlyTag("v");
						if(!v){
							v = accessor().createElement("v");
							node.appendChild(v);
						}
						v.textContent = val;
					}else if(typeof val == "string" && val){
						node.attr("t", "s");
						var n = _.findIndex(SST.childNodes, function(si){
							return si.onlyTag("t").textContent == val;
						});
						if(n > 0){
							node.onlyTag("v").textContent = n;
						}else{
							node.onlyTag("v").textContent = SST.childNodes.length;
							var si = accessor().createElement("si");
							var t = accessor().createElement("t");
							t.textContent = val;
							SST.appendChild(si);
							si.appendChild(t);
						}
					}else if(_.isNull(val)){
						rowNode.removeChild(node);
					}else if(val === ""){
						var v = node.onlyTag("v");
						if(v){
							node.removeChild(v);
						}
					}else if(_.isUndefined(val)){
						//do nothing. twice
					}else{
						throw new Error("Unsupported cell value");
					}

				})
			})
		}else{
			return rows.map(function(r){
				var rowNode = rowNodes.find({r:r});
				if(!rowNode){
					return columns.map(function(){return null});
				}else{
					var cellNodes = rowNode.byTag("c");
					return columns.map(function(c){
						var node = cellNodes.find({r: addr(c, r)});
						if(!node){
							return null;
						}
						var vnode = node.onlyTag("v");
						if(!vnode){
							return "";
						}
						var val = vnode.textContent;
						if(node.attr("t") == "s"){
							val = SST.childNodes[val].onlyTag("t").textContent;
						}
						return val;
					});
				}
			})
		}
	}
}

module.exports = Workbook;