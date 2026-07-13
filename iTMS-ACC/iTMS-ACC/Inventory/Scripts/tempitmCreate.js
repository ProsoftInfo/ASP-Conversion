(function (window, document) {
	"use strict";

	var statusText = "";
	var tempCode = "";
	var tempDesc = "";
	var tempShortDesc = "";
	var tempAddDesc = "";
	var tempItemCode = "";
	var returnText = "-1";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function field(name) {
		var frm = form();
		return frm && frm.elements ? frm.elements[name] || null : null;
	}

	function valueOf(name) {
		var item = field(name);
		return item ? item.value : "";
	}

	function setValue(name, value) {
		var item = field(name);
		if (item) {
			item.value = value == null ? "" : String(value);
		}
	}

	function currentItemType() {
		return trim(valueOf("hItemType") || window.iType || "");
	}

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function xmlObject(name) {
		ensureCompat();
		return window[name] || document[name] || document.getElementById(name) || null;
	}

	function xmlRoot(nameOrObject) {
		var object = typeof nameOrObject === "string" ? xmlObject(nameOrObject) : nameOrObject;
		if (window.ITMSModalReturnCompat && window.ITMSModalReturnCompat.xmlRoot) {
			return window.ITMSModalReturnCompat.xmlRoot(object);
		}
		return object && object.documentElement || object && object.XMLDocument && object.XMLDocument.documentElement || object && object._doc && object._doc.documentElement || object && object.nodeType === 1 && object || null;
	}

	function xmlDocument(name) {
		var object = xmlObject(name);
		var root = xmlRoot(object);
		return object && object.XMLDocument || object && object._doc || root && root.ownerDocument || null;
	}

	function childElements(node) {
		var result = [];
		for (var i = 0; node && i < node.childNodes.length; i += 1) {
			if (node.childNodes[i].nodeType === 1) {
				result.push(node.childNodes[i]);
			}
		}
		return result;
	}

	function attrByIndex(node, index) {
		return node && node.attributes && node.attributes[index] ? node.attributes[index].value : "";
	}

	function attrNameByIndex(node, index) {
		return node && node.attributes && node.attributes[index] ? node.attributes[index].name : "";
	}

	function serializeXml(rootOrDoc) {
		var doc = rootOrDoc && rootOrDoc.nodeType === 9 ? rootOrDoc : rootOrDoc && rootOrDoc.ownerDocument;
		return new XMLSerializer().serializeToString(doc || rootOrDoc);
	}

	function loadXmlIntoIsland(name, url) {
		var xhr = new XMLHttpRequest();
		var object;
		var text;
		xhr.open("GET", url, false);
		xhr.send(null);
		text = xhr.responseText || "";
		if (!text && xhr.responseXML) {
			text = new XMLSerializer().serializeToString(xhr.responseXML);
		}
		if (text) {
			object = xmlObject(name);
			if (object && object.loadXML) {
				object.loadXML(text);
			} else if (object && object.LoadXML) {
				object.LoadXML(text);
			}
		}
		return text ? xmlRoot(name) : null;
	}

	function returnAndClose(value) {
		returnText = value;
		if (window.ITMSModalReturnCompat && window.ITMSModalReturnCompat.returnAndClose) {
			window.ITMSModalReturnCompat.returnAndClose(value);
		} else {
			window.close();
		}
	}

	function trimTrue(value) {
		return trim(value);
	}

	function LetIType(obj) {
		var type = obj ? obj.value : "";
		window.iType = type;
		setValue("hItemType", type);
		setValue("txtItmDesc", "");
		setValue("itmCode", "");
		if (type === "YRN") {
			field("itmCode").readOnly = true;
			field("btnYrnCode").disabled = false;
		} else {
			field("itmCode").readOnly = false;
			field("btnYrnCode").disabled = true;
		}
	}

	function CreateItemCode() {
		var opener = window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog;
		var url = "itmCodeCreate.asp";
		if (opener) {
			opener(url, "", "dialogHeight:270px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No", function (value) {
				var parts = String(value || "").split("``");
				setValue("txtItmDesc", parts[0] || "");
				setValue("itmCode", parts[1] || "");
				if (field("itmCode")) {
					field("itmCode").readOnly = true;
				}
			});
		} else {
			window.open(url, "_blank", "height=270,width=450,resizable=no,status=no");
		}
		return false;
	}

	function itemCodeCheck(value) {
		var root;
		var exists = false;
		var bClaExists = false;
		var bItmExists = false;
		statusText = "";
		tempCode = "";
		tempDesc = "";
		tempShortDesc = "";
		tempAddDesc = "";
		tempItemCode = "";
		root = loadXmlIntoIsland("Data", "itmCodeXMLSelect.asp?sOrgCode=" + encodeURIComponent(trim(valueOf("hOrgCode"))));
		childElements(root).forEach(function (node) {
			var code = attrByIndex(node, 0);
			var firstName;
			if (exists || trim(String(value)).toLowerCase() !== trim(code).toLowerCase()) {
				return;
			}
			firstName = attrNameByIndex(node, 0);
			if (/^T/i.test(firstName)) {
				tempCode = code;
				tempDesc = attrByIndex(node, 2);
				tempShortDesc = attrByIndex(node, 1);
				tempAddDesc = attrByIndex(node, 3);
				tempItemCode = attrByIndex(node, 4);
				if (attrByIndex(node, 5) === "N") {
					statusText = "Temporary Item code already Exists";
				} else if (attrByIndex(node, 5) === "Y") {
					statusText = " Temporary Item code already been created / mapped with Permanent Item \n [" + attrByIndex(node, 6) + "]. So select this Item from the catalogue itself.";
				}
			} else if (/^C/i.test(firstName)) {
				bClaExists = true;
			} else if (/^I/i.test(firstName)) {
				bItmExists = true;
			}
			if (bItmExists) {
				statusText = "Item code already Exists but not defined for the Unit";
			} else if (bClaExists) {
				statusText = "Item code already Exists";
			}
			exists = true;
		});
		return exists;
	}

	function InsertDetails(itemCode, itemDesc, itemShortDesc, itemAddDesc) {
		var root = xmlRoot("OutData");
		var doc = xmlDocument("OutData");
		var node;
		if (!root || !doc) {
			return null;
		}
		root.setAttribute("ITMTYPE", "");
		root.setAttribute("APPCODE", trim(valueOf("hAppCode")));
		root.setAttribute("MODCODE", trim(valueOf("hModCode")));
		root.setAttribute("CRESTAGE", trim(valueOf("hCreStage")));
		node = doc.createElement("TDETAILS");
		node.setAttribute("ITMCODE", itemCode);
		node.setAttribute("ITMDESC", itemDesc);
		node.setAttribute("ITMSHDESC", itemShortDesc);
		node.setAttribute("ITMADDDESC", itemAddDesc);
		root.appendChild(node);
		return root;
	}

	function acceptExistingTemp() {
		setValue("itmCode", tempCode);
		setValue("txtItmDesc", tempDesc);
		setValue("txtItmShDesc", tempShortDesc);
		setValue("txtItmAddDesc", tempAddDesc);
		InsertDetails(trim(valueOf("itmCode")), trim(valueOf("txtItmDesc")), trim(valueOf("txtItmShDesc")), trim(valueOf("txtItmAddDesc")));
		tempItemCode = tempItemCode || "";
		returnAndClose(trim(valueOf("itmCode")) + "``" + trim(valueOf("txtItmShDesc")));
	}

	function validateExistingCode(type) {
		var code = trim(valueOf("itmCode"));
		if (type !== "YRN" && !code) {
			alert("Enter Item Code");
			field("itmCode").select();
			return false;
		}
		if (code && itemCodeCheck(code)) {
			if (/^T/.test(statusText)) {
				if (confirm(statusText + ". Do you want to have the existing Temporary Item Code?")) {
					acceptExistingTemp();
				} else {
					field("itmCode").select();
				}
				return false;
			}
			alert(statusText);
			field("itmCode").select();
			return false;
		}
		return true;
	}

	function CheckSubmit() {
		var type = currentItemType();
		var root;
		var xhr;
		if (!validateExistingCode(type)) {
			return false;
		}
		if (!trim(valueOf("txtItmDesc"))) {
			alert("Enter Item Description");
			field("txtItmDesc").select();
			return false;
		}
		if (!trim(valueOf("txtItmShDesc"))) {
			alert("Enter Item Short Description");
			field("txtItmShDesc").select();
			return false;
		}
		root = InsertDetails(trim(valueOf("itmCode")), trim(valueOf("txtItmDesc")), trim(valueOf("txtItmShDesc")), trim(valueOf("txtItmAddDesc")));
		xhr = new XMLHttpRequest();
		xhr.open("POST", "tempitmInsert.asp", false);
		xhr.send(root ? serializeXml(root) : "");
		if (!xhr.responseText) {
			returnAndClose(trim(valueOf("itmCode")) + "``" + trim(valueOf("txtItmShDesc")));
		} else {
			alert(xhr.responseText);
		}
		return false;
	}

	if (window.ITMSModalReturnCompat && window.ITMSModalReturnCompat.install) {
		window.ITMSModalReturnCompat.install(function () {
			return returnText === "YES" ? trim(valueOf("itmCode")) + "``" + trim(valueOf("txtItmShDesc")) : returnText;
		});
	}

	window.trimTrue = trimTrue;
	window.LetIType = LetIType;
	window.CreateItemCode = CreateItemCode;
	window.itemCodeCheck = itemCodeCheck;
	window.CheckSubmit = CheckSubmit;
	window.InsertDetails = InsertDetails;
}(window, document));
