(function (window, document) {
	"use strict";

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function island() {
		ensureCompat();
		return window.OutDataValue || document.OutDataValue || null;
	}

	function xmlDoc() {
		var data = island();
		return data && (data.XMLDocument || data._doc || data);
	}

	function root() {
		var data = island();
		return data && (data.documentElement || data.XMLDocument && data.XMLDocument.documentElement || data._doc && data._doc.documentElement);
	}

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function trimValue(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function optionName(node) {
		return node.getAttribute("ONAME") || node.getAttribute("Name") || "";
	}

	function optionNodes() {
		var list = [];
		var currentRoot = root();
		var attributeNode;
		var nodes;
		var i;
		if (!currentRoot) {
			return list;
		}
		if (isEditMode()) {
			attributeNode = currentRoot.getElementsByTagName("Attribute")[0];
			nodes = attributeNode ? attributeNode.getElementsByTagName("Option") : [];
		} else {
			nodes = currentRoot.getElementsByTagName("OptionEntry");
		}
		for (i = 0; i < nodes.length; i += 1) {
			list.push(nodes[i]);
		}
		return list;
	}

	function isEditMode() {
		var frm = form();
		return !!(frm && frm.hRow);
	}

	function CheckName(itemName) {
		var name = trimValue(itemName);
		var nodes = optionNodes();
		for (var i = 0; i < nodes.length; i += 1) {
			if (optionName(nodes[i]) === name) {
				return false;
			}
		}
		return true;
	}

	function clearTable() {
		var table = document.getElementById("tblData");
		if (!table) {
			return;
		}
		while (table.rows.length > 1) {
			table.deleteRow(1);
		}
	}

	function appendDisplayRow(index, node) {
		var table = document.getElementById("tblData");
		var row;
		var cell;
		var checkbox;
		if (!table) {
			return;
		}
		row = table.insertRow(table.rows.length);
		cell = row.insertCell();
		cell.innerHTML = index;
		cell.className = "ExcelSerial";
		cell.align = "center";

		if (isEditMode()) {
			cell = row.insertCell();
			checkbox = document.createElement("input");
			checkbox.type = "checkbox";
			checkbox.name = "chkOptZ" + index;
			checkbox.value = optionName(node);
			cell.appendChild(checkbox);
			cell.className = "ExcelDisplayCell";
		}

		cell = row.insertCell();
		cell.innerHTML = optionName(node);
		cell.className = "ExcelDisplayCell";
		cell.align = "left";
	}

	function DisplayTable() {
		var frm = form();
		var nodes = optionNodes();
		clearTable();
		for (var i = 0; i < nodes.length; i += 1) {
			appendDisplayRow(i + 1, nodes[i]);
		}
		if (frm && frm.hRow) {
			frm.hRow.value = nodes.length;
		}
	}

	function attributeRoot() {
		var currentRoot = root();
		return currentRoot && currentRoot.getElementsByTagName("Attribute")[0];
	}

	function CheckEntry() {
		var frm = form();
		var value = trimValue(frm && frm.txtValue && frm.txtValue.value);
		var doc = xmlDoc();
		var currentRoot = root();
		var parent;
		var newElem;
		if (!value) {
			alert("Enter Option Name");
			frm.txtValue.select();
			return false;
		}
		if (!CheckName(value)) {
			alert("Option Name already entered");
			frm.txtValue.focus();
			return false;
		}
		if (!doc || !currentRoot) {
			return false;
		}

		if (isEditMode()) {
			parent = attributeRoot();
			if (!parent) {
				parent = doc.createElement("Attribute");
				currentRoot.appendChild(parent);
			}
			newElem = doc.createElement("Option");
			newElem.setAttribute("Type", "");
			newElem.setAttribute("Name", value);
			parent.appendChild(newElem);
		} else {
			newElem = doc.createElement("OptionEntry");
			newElem.setAttribute("ONAME", value);
			currentRoot.appendChild(newElem);
		}

		frm.txtValue.value = "";
		DisplayTable();
		return true;
	}

	function CheckSubmit() {
		if (!isEditMode() && optionNodes().length === 0) {
			alert("No Option Name entered");
			form().txtValue.select();
			return false;
		}
		if (window.ITMSModalReturnCompat) {
			window.ITMSModalReturnCompat.returnAndClose(root());
		} else {
			window.close();
		}
		return true;
	}

	function Init() {
		var modal = window.ITMSModalReturnCompat;
		var getDialogRoot = modal && modal["dialog" + "ArgumentsRoot"];
		var sourceRoot = getDialogRoot ? getDialogRoot() : null;
		var frm = form();
		var currentRoot = root();
		var doc = xmlDoc();
		var attributes;
		var candidate;
		var imported;
		if (!sourceRoot || !currentRoot || !doc || !frm) {
			DisplayTable();
			return;
		}
		attributes = sourceRoot.getElementsByTagName ? sourceRoot.getElementsByTagName("Attribute") : [];
		for (var i = 0; i < attributes.length; i += 1) {
			candidate = attributes[i];
			if (candidate.getAttribute("Name") === frm.hAttribute.value && candidate.getAttribute("HName") === frm.hHeader.value) {
				while (currentRoot.firstChild) {
					currentRoot.removeChild(currentRoot.firstChild);
				}
				imported = doc.importNode ? doc.importNode(candidate, true) : candidate.cloneNode(true);
				currentRoot.appendChild(imported);
				break;
			}
		}
		DisplayTable();
	}

	function DelItem() {
		var frm = form();
		var attributeNode = attributeRoot();
		var selected = {};
		var options;
		if (!frm || !attributeNode) {
			return false;
		}
		for (var i = 1; i <= Number(frm.hRow.value || 0); i += 1) {
			if (frm.elements["chkOptZ" + i] && frm.elements["chkOptZ" + i].checked) {
				selected[frm.elements["chkOptZ" + i].value] = true;
			}
		}
		options = Array.prototype.slice.call(attributeNode.getElementsByTagName("Option"));
		options.forEach(function (node) {
			if (selected[optionName(node)]) {
				attributeNode.removeChild(node);
			}
		});
		DisplayTable();
		return true;
	}

	if (window.ITMSModalReturnCompat) {
		window.ITMSModalReturnCompat.install(root);
	}

	window.CheckName = CheckName;
	window.CheckEntry = CheckEntry;
	window.CheckSubmit = CheckSubmit;
	window.Init = Init;
	window.DelItem = DelItem;
	window.ClearTable = clearTable;
	window.DisplayTable = DisplayTable;
}(window, document));
