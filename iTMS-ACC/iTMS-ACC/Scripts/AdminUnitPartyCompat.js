(function (window, document) {
	"use strict";

	var xmlDoc = null;
	var unitNodes = [];

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.formname || document.forms.formname || document.forms[0] || null;
	}

	function createXmlDocument(rootName) {
		return document.implementation.createDocument("", rootName, null);
	}

	function parseXml(text) {
		var parsed;
		if (!trim(text)) {
			return createXmlDocument("ROOT");
		}
		parsed = new DOMParser().parseFromString(text, "application/xml");
		if (parsed.getElementsByTagName("parsererror").length) {
			return createXmlDocument("ROOT");
		}
		return parsed;
	}

	function serializeXml(doc) {
		return new XMLSerializer().serializeToString(doc);
	}

	function attr(node, name) {
		return node && node.getAttribute ? trim(node.getAttribute(name)) : "";
	}

	function setAttr(node, name, value) {
		if (node && node.setAttribute) {
			node.setAttribute(name, value == null ? "" : String(value));
		}
	}

	function childElements(node) {
		var result = [];
		var children = node ? node.childNodes : [];
		var i;
		for (i = 0; i < children.length; i += 1) {
			if (children[i].nodeType === 1) {
				result.push(children[i]);
			}
		}
		return result;
	}

	function table() {
		return document.getElementById("tblDetails");
	}

	function clearDetailRows() {
		var detailTable = table();
		while (detailTable && detailTable.rows.length > 2) {
			detailTable.deleteRow(2);
		}
	}

	function appendCell(row, text, className, align, colspan) {
		var cell = row.insertCell(-1);
		cell.className = className || "";
		cell.align = align || "";
		if (colspan) {
			cell.colSpan = colspan;
		}
		if (text != null) {
			cell.appendChild(document.createTextNode(String(text)));
		}
		return cell;
	}

	function addHidden(name, value) {
		var frm = form();
		var input = document.createElement("input");
		input.type = "hidden";
		input.name = name;
		input.value = value || "";
		if (frm) {
			frm.appendChild(input);
		}
		return input;
	}

	function addPartyInput(name, value) {
		var input = document.createElement("input");
		input.type = "text";
		input.name = name;
		input.value = value || "";
		input.className = "FormElemRead";
		input.readOnly = true;
		input.size = 60;
		return input;
	}

	function addPartyButton(index, label) {
		var button = document.createElement("input");
		button.type = "button";
		button.name = "Btn:" + index;
		button.value = label;
		button.className = "AddButtonX";
		button.onclick = function () {
			return window.GetParty(button);
		};
		return button;
	}

	function renderDetails() {
		var root = xmlDoc && xmlDoc.documentElement;
		var detailTable = table();
		var headers;
		var headerIndex;
		var unitIndex = 0;
		var rows;
		var row;
		var cell;
		var node;
		if (!detailTable || !root) {
			return;
		}

		clearDetailRows();
		unitNodes = [];
		headers = childElements(root);
		for (headerIndex = 0; headerIndex < headers.length; headerIndex += 1) {
			row = detailTable.insertRow(-1);
			appendCell(row, attr(headers[headerIndex], "ORGANIZATIONNAME"), "ExcelDisplayCell", "left", 4);
			rows = childElements(headers[headerIndex]);
			for (var rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
				node = rows[rowIndex];
				unitIndex += 1;
				unitNodes[unitIndex] = node;

				row = detailTable.insertRow(-1);
				appendCell(row, unitIndex, "ExcelSerial", "center");
				appendCell(row, attr(node, "ORGANIZATIONNAME"), "ExcelDisplayCell", "left");
				addHidden("txtOrg" + unitIndex, attr(node, "ORGANIZATIONID"));

				cell = appendCell(row, "", "ExcelDisplayCell", "left");
				cell.appendChild(addPartyInput("txtParty" + unitIndex, attr(node, "PARTYNAME")));

				cell = appendCell(row, "", "ExcelDisplayCell", "center");
				cell.appendChild(addPartyButton(unitIndex, attr(node, "PARTYNAME") ? "Change" : "Allocate"));
			}
		}
	}

	function requestXml(url, callback) {
		var xhr = new XMLHttpRequest();
		xhr.open("GET", url, true);
		xhr.onreadystatechange = function () {
			if (xhr.readyState !== 4) {
				return;
			}
			if (xhr.status >= 200 && xhr.status < 300) {
				callback(xhr.responseXML && xhr.responseXML.documentElement ? xhr.responseXML : parseXml(xhr.responseText));
			} else {
				alert("Unable to load unit party details.");
			}
		};
		xhr.send(null);
	}

	function postXml(url, doc, callback) {
		var xhr = new XMLHttpRequest();
		xhr.open("POST", url, true);
		xhr.setRequestHeader("Content-Type", "text/xml; charset=UTF-8");
		xhr.onreadystatechange = function () {
			if (xhr.readyState !== 4) {
				return;
			}
			if (xhr.status >= 200 && xhr.status < 300) {
				callback(xhr.responseText || "");
			} else {
				alert("Unable to save unit party details.");
			}
		};
		xhr.send(serializeXml(doc));
	}

	function openDialog(url, features, callback) {
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog(url, "", features || "", callback || function () {});
			return true;
		}
		alert("The compatibility script is still loading. Please try again.");
		return false;
	}

	function runPartyDialog(query, callback) {
		var url = "PartySelection.asp" + (query ? "?" + query : "");
		openDialog(url, "dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No", function (outValue) {
			var text = trim(outValue);
			var parts;
			if (!text) {
				return;
			}
			parts = text.split(":");
			if (parts.length === 1) {
				runPartyDialog(text, callback);
				return;
			}
			callback(parts);
		});
	}

	function controlByName(name) {
		var frm = form();
		return frm && frm.elements ? frm.elements[name] : null;
	}

	window.Init = function () {
		requestXml("XMLSelectUnitParty.asp", function (doc) {
			xmlDoc = doc;
			renderDetails();
		});
	};

	window.GetParty = function (obj) {
		var index = String(obj && obj.name || "").split(":")[1];
		var partyField;
		var node = unitNodes[index];
		if (!index || !node) {
			return false;
		}

		runPartyDialog("", function (parts) {
			if (trim(parts[0]) === "-1") {
				return;
			}
			partyField = controlByName("txtParty" + index);
			if (partyField) {
				partyField.value = parts[0] || "";
			}
			setAttr(node, "FLAG", "Y");
			setAttr(node, "PARTYCODE", parts[1] || "");
			setAttr(node, "PARTYNAME", parts[0] || "");
		});
		return false;
	};

	window.CheckSubmit = function () {
		if (!xmlDoc) {
			alert("Unit party details are still loading. Please try again.");
			return false;
		}
		postXml("UnitPartyInsert.asp", xmlDoc, function (responseText) {
			if (trim(responseText) === "") {
				window.location.href = "../welcome_admin.asp";
			} else {
				alert(responseText);
			}
		});
		return false;
	};
}(window, document));
