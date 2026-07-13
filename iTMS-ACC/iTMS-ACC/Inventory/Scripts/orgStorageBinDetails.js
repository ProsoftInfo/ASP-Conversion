(function (window, document) {
	"use strict";

	var objTemp = window["dialog" + "Arguments"] || null;

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function field(name) {
		var frm = form();
		return document.getElementById(name) || (frm && frm.elements ? frm.elements[name] : null);
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

	function xmlDoc() {
		return objTemp && objTemp.XMLDocument || objTemp && objTemp._doc || objTemp && objTemp.ownerDocument || null;
	}

	function rootNode() {
		return objTemp && objTemp.documentElement || objTemp && objTemp.XMLDocument && objTemp.XMLDocument.documentElement || objTemp && objTemp._doc && objTemp._doc.documentElement || objTemp || null;
	}

	function selectNodes(node, expression) {
		return node && node.selectNodes ? node.selectNodes(expression) : [];
	}

	function serializeXml(nodeOrDoc) {
		var target = nodeOrDoc && nodeOrDoc.nodeType === 9 ? nodeOrDoc : nodeOrDoc && nodeOrDoc.ownerDocument || nodeOrDoc;
		return target ? new XMLSerializer().serializeToString(target) : "";
	}

	function saveXml() {
		var xhr = new XMLHttpRequest();
		xhr.open("POST", "XMLSave.asp?Name=StorageNew", false);
		xhr.send(serializeXml(xmlDoc() || rootNode()));
	}

	function returnDialogValue() {
		var root = rootNode();
		if (window.ITMSModalReturnCompat && window.ITMSModalReturnCompat.returnValue) {
			window.ITMSModalReturnCompat.returnValue(root);
		} else if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(root);
		}
	}

	function CheckSubmit() {
		var root = rootNode();
		var storage = selectNodes(root, "//Organization/Storage")[0];
		var bins = selectNodes(root, "//Organization/Storage/Bin");
		var count = Number(valueOf("hBinNo")) || 0;
		var bin;
		if (storage) {
			storage.setAttribute("NUMBEROFBINS", String(count));
		}
		for (var i = 1; i <= bins.length; i += 1) {
			bin = bins[i - 1];
			bin.setAttribute("BINNUMBER", valueOf("Chk" + i));
			bin.setAttribute("BINCODE", valueOf("txtBinCode" + i));
			bin.setAttribute("BINNAME", valueOf("txtBinName" + i));
			bin.setAttribute("BINAREA", valueOf("txtBinSize" + i));
		}
		saveXml();
		returnDialogValue();
		window.close();
		return false;
	}

	function DeleteItems() {
		var root = rootNode();
		var bins = selectNodes(root, "//Organization/Storage/Bin");
		var selected = {};
		var changed = false;
		var box;
		for (var i = 1; i <= Number(valueOf("hBinNo") || 0); i += 1) {
			box = field("Chk" + i);
			if (box && box.checked) {
				selected[trim(box.value)] = true;
			}
		}
		for (var j = bins.length - 1; j >= 0; j -= 1) {
			if (selected[trim(bins[j].getAttribute("SLNO"))] && bins[j].parentNode) {
				bins[j].parentNode.removeChild(bins[j]);
				changed = true;
			}
		}
		if (changed) {
			bins = selectNodes(root, "//Organization/Storage/Bin");
			for (var k = 0; k < bins.length; k += 1) {
				bins[k].setAttribute("SLNO", String(k + 1));
			}
			setValue("hBinNo", bins.length);
		}
		DisplaytableBin();
		return false;
	}

	function AddNew() {
		var root = rootNode();
		var doc = xmlDoc();
		var count = Number(valueOf("hBinNo")) + 1;
		var storages = selectNodes(root, "//Organization/Storage");
		var bin;
		if (!doc) {
			return false;
		}
		for (var i = 0; i < storages.length; i += 1) {
			storages[i].setAttribute("NUMBEROFBINS", String(count));
			bin = doc.createElement("Bin");
			bin.setAttribute("SLNO", String(count));
			bin.setAttribute("BINNUMBER", "");
			bin.setAttribute("BINCODE", "");
			bin.setAttribute("BINNAME", "");
			bin.setAttribute("BINAREA", "");
			storages[i].appendChild(bin);
		}
		setValue("hBinNo", count);
		DisplaytableBin();
		return false;
	}

	function clearTable() {
		var table = document.getElementById("tblBin");
		while (table && table.rows.length) {
			table.deleteRow(0);
		}
	}

	function appendCell(row, value, className, align) {
		var cell = row.insertCell();
		cell.textContent = value == null ? "" : String(value).replace(/&nbsp;/g, " ");
		cell.className = className || "";
		if (align) {
			cell.align = align;
		}
		return cell;
	}

	function appendInput(row, type, name, value, size, maxLength) {
		var cell = row.insertCell();
		var input = document.createElement("input");
		input.type = type;
		input.name = name;
		input.value = value == null ? "" : String(value);
		input.className = "Formelem";
		if (size) {
			input.size = size;
		}
		if (maxLength) {
			input.maxLength = maxLength;
		}
		cell.appendChild(input);
		cell.className = type === "checkbox" ? "ExcelDisplayCell" : "ExcelInputCell";
		return cell;
	}

	function DisplaytableBin() {
		var root = rootNode();
		var table = document.getElementById("tblBin");
		var bins = selectNodes(root, "//Organization/Storage/Bin");
		var row;
		var deleteLink;
		var deleteImage;
		var bin;
		clearTable();
		row = table.insertRow(0);
		appendCell(row, "S.No.", "ExcelHeaderCell", "center");
		deleteLink = document.createElement("a");
		deleteLink.href = "#";
		deleteLink.onclick = function () {
			DeleteItems();
			return false;
		};
		deleteImage = document.createElement("img");
		deleteImage.border = 0;
		deleteImage.src = "../../assets/images/iTMS%20Icons/DeleteIcon.gif";
		deleteImage.width = 15;
		deleteImage.height = 15;
		deleteLink.appendChild(deleteImage);
		appendCell(row, "", "ExcelHeaderCell", "center").appendChild(deleteLink);
		appendCell(row, "Bin Code", "ExcelHeaderCell", "center");
		appendCell(row, "Bin Description", "ExcelHeaderCell", "center");
		appendCell(row, "&nbsp;Bin Size / Area", "ExcelHeaderCell", "center");
		for (var i = 0; i < bins.length; i += 1) {
			bin = bins[i];
			row = table.insertRow(table.rows.length);
			appendCell(row, i + 1, "ExcelDisplayCell", "center");
			appendInput(row, "checkbox", "Chk" + (i + 1), String(i + 1), 0, 0);
			appendInput(row, "text", "txtBinCode" + (i + 1), bin.getAttribute("BINNUMBER") || "", 12, 10);
			appendInput(row, "text", "txtBinName" + (i + 1), bin.getAttribute("BINNAME") || "", 50, 40);
			appendInput(row, "text", "txtBinSize" + (i + 1), bin.getAttribute("BINAREA") || "", 12, 10);
		}
	}

	function ClearAll() {
		var select = field("selLocName");
		if (select) {
			select.options.length = 0;
			select.options[0] = new Option("Select", "select");
		}
		clearTable();
		return false;
	}

	function FnInit() {
		var root = rootNode();
		var organizations = root ? root.childNodes : [];
		var org;
		var storage;
		var count;
		var doc = xmlDoc();
		var bin;
		for (var i = 0; i < organizations.length; i += 1) {
			org = organizations[i];
			if (org.nodeType !== 1 || trim(org.nodeName) !== "Organization") {
				continue;
			}
			for (var j = 0; j < org.childNodes.length; j += 1) {
				storage = org.childNodes[j];
				if (storage.nodeType !== 1 || trim(storage.nodeName) !== "Storage") {
					continue;
				}
				count = Number(storage.getAttribute("NUMBEROFBINS")) || 0;
				setValue("hBinNo", count);
				if (document.getElementById("UnitID")) {
					document.getElementById("UnitID").textContent = org.getAttribute("UNITNAME") || "";
				}
				if (document.getElementById("LocID")) {
					document.getElementById("LocID").textContent = storage.getAttribute("LOCATIONNAME") || "";
				}
				setValue("hLocNo", storage.getAttribute("LOCATIONNUMBER") || "");
				if (!storage.hasChildNodes()) {
					for (var k = 1; k <= count; k += 1) {
						bin = doc.createElement("Bin");
						bin.setAttribute("SLNO", String(k));
						bin.setAttribute("BINNUMBER", "");
						bin.setAttribute("BINCODE", "");
						bin.setAttribute("BINNAME", "");
						bin.setAttribute("BINAREA", "");
						storage.appendChild(bin);
					}
				}
				DisplaytableBin();
				return false;
			}
		}
		DisplaytableBin();
		return false;
	}

	window.CheckSubmit = CheckSubmit;
	window.checkSubmit = CheckSubmit;
	window.window_Unload = returnDialogValue;
	window.DeleteItems = DeleteItems;
	window.AddNew = AddNew;
	window.DisplaytableBin = DisplaytableBin;
	window.ClearTable = clearTable;
	window.ClearAll = ClearAll;
	window.FnInit = FnInit;
	window.addEventListener("beforeunload", returnDialogValue);
}(window, document));
