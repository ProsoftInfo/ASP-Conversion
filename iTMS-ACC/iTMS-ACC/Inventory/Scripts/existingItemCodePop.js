(function (window, document) {
	"use strict";

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

	function selectedText(name) {
		var select = field(name);
		return select && select.selectedIndex >= 0 && select.options[select.selectedIndex] ? select.options[select.selectedIndex].text : "";
	}

	function ClearTable() {
		var table = document.getElementById("tblDetails");
		while (table && table.rows.length > 1) {
			table.deleteRow(1);
		}
	}

	function DisableTxtQty(obj) {
		var frm = form();
		ClearTable();
		frm.selSearchBy.selectedIndex = 0;
		frm.selSearchBy.disabled = true;
		frm.btnSearch.disabled = true;
		frm.txtSearchFor.readOnly = true;
		if (obj && obj.value === "I") {
			frm.btnSearch.disabled = false;
			frm.txtSearchFor.readOnly = false;
			frm.selSearchBy.disabled = false;
		} else if (obj && obj.value === "A") {
			SearchItem();
		}
		return false;
	}

	function xmlFromResponse(xhr) {
		if (xhr.responseXML && xhr.responseXML.documentElement) {
			return xhr.responseXML;
		}
		return new DOMParser().parseFromString(xhr.responseText || "<root/>", "application/xml");
	}

	function appendCell(row, value, className, align) {
		var cell = row.insertCell();
		cell.textContent = trim(value);
		cell.className = className || "ExcelDisplayCell";
		cell.align = align || "left";
		return cell;
	}

	function attr(node, index) {
		return node && node.attributes && node.attributes[index] ? node.attributes[index].nodeValue : "";
	}

	function renderRows(root, itemType) {
		var table = document.getElementById("tblDetails");
		var nodes = root ? root.childNodes : [];
		var row;
		var cell;
		var link;
		var itemCode;
		var serial = 0;
		for (var i = 0; i < nodes.length; i += 1) {
			if (nodes[i].nodeType !== 1) {
				continue;
			}
			serial += 1;
			itemCode = attr(nodes[i], 0);
			row = table.insertRow(table.rows.length);
			appendCell(row, serial, "ExcelSerial", "center");
			cell = row.insertCell();
			cell.className = "ExcelDisplayCell";
			cell.align = "left";
			if (itemType === "YRN" || itemType === "FAB") {
				link = document.createElement("a");
				link.href = "#";
				link.name = String(serial);
				link.className = "ExcelDisplayLink";
				link.textContent = trim(itemCode);
				link.onclick = function (code) {
					return function () {
						DisplayItemCode(code);
						return false;
					};
				}(itemCode);
				cell.appendChild(link);
			} else {
				cell.textContent = trim(itemCode);
			}
			appendCell(row, attr(nodes[i], 1));
			appendCell(row, attr(nodes[i], 2));
			appendCell(row, attr(nodes[i], 3));
		}
		return serial;
	}

	function SearchItem() {
		var itemType = valueOf("hItmType");
		var searchBy = field("selSearchBy");
		var searchFor = field("txtSearchFor");
		var who;
		var search;
		var xhr;
		var xml;
		ClearTable();
		if (!field("radDisplay") || !field("radDisplay").checked) {
			if (searchBy.selectedIndex === 0) {
				alert("Select Search By");
				searchBy.focus();
				return false;
			}
			if (trim(searchFor.value) === "") {
				alert("Enter " + selectedText("selSearchBy") + " to Search For");
				searchFor.focus();
				return false;
			}
			who = trim(searchBy.value);
			search = trim(searchFor.value);
		} else {
			who = "ALL";
			search = "";
		}
		xhr = new XMLHttpRequest();
		xhr.open("GET", "XMLItemCode.asp?sIType=" + encodeURIComponent(itemType) + "&sWho=" + encodeURIComponent(who) + "&sSearch=" + encodeURIComponent(search), false);
		xhr.send();
		if (trim(xhr.responseText) === "") {
			alert(xhr.responseText);
			return false;
		}
		xml = xmlFromResponse(xhr);
		if (!xml.documentElement || renderRows(xml.documentElement, itemType) === 0) {
			alert("No Item matching your Search criteria.");
		}
		return false;
	}

	function DisplayItemCode(itemCode) {
		var itemType = valueOf("hItmType");
		var height = itemType === "FAB" ? "410" : "250";
		var url = "itmCodeDisplay.asp?ItemCode=" + encodeURIComponent(itemCode) + "&ItemType=" + encodeURIComponent(itemType);
		if (itemType !== "YRN" && itemType !== "FAB") {
			return false;
		}
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog(url, "", "dialogHeight:" + height + "px;dialogWidth:400px;center:Yes;help:No;resizable:No;status:No");
		} else {
			window.open(url, "_blank", "height=" + height + ",width=400,resizable=no,status=no,scrollbars=yes");
		}
		return false;
	}

	window.DisableTxtQty = DisableTxtQty;
	window.SearchItem = SearchItem;
	window.ClearTable = ClearTable;
	window.DisplayItemCode = DisplayItemCode;
}(window, document));
