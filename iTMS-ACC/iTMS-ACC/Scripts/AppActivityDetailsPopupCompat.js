(function (window, document) {
	"use strict";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.formname || document.forms.formname || document.forms[0] || null;
	}

	function field(name) {
		var frm = form();
		return frm && (frm.elements[name] || frm[name]) || document.getElementById(name) || null;
	}

	function value(name) {
		var item = field(name);
		return item ? item.value : "";
	}

	function setValue(name, data) {
		var item = field(name);
		if (item) {
			item.value = data == null ? "" : String(data);
		}
	}

	function upgradeXml() {
		if (window.ITMSModernCompat) {
			window.ITMSModernCompat.upgradeXmlIslands(document);
		}
	}

	function xmlObject(name) {
		var element;
		upgradeXml();
		element = document.getElementById(name);
		return window[name] || document[name] || element && element._itmsXmlIsland || element || null;
	}

	function xmlDocument(name) {
		var data = xmlObject(name);
		return data && (data.XMLDocument || data._doc || data) || null;
	}

	function xmlRoot(name) {
		var data = xmlObject(name);
		var doc = xmlDocument(name);
		return data && data.documentElement || doc && doc.documentElement || null;
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

	function detailNodes() {
		return childElements(xmlRoot("OutData")).filter(function (node) {
			return node.nodeName === "DETAILS";
		});
	}

	function getAttr(node, name) {
		return node && node.getAttribute(name) || "";
	}

	function setAttr(node, name, data) {
		if (node) {
			node.setAttribute(name, data == null ? "" : String(data));
		}
	}

	function serializeXml(name) {
		var doc = xmlDocument(name);
		if (doc && typeof XMLSerializer !== "undefined") {
			return new XMLSerializer().serializeToString(doc);
		}
		return doc && doc.xml || "";
	}

	function postXml(url, islandName) {
		var xhr = new XMLHttpRequest();
		xhr.open("POST", url, false);
		xhr.setRequestHeader("Content-Type", "text/xml");
		xhr.send(islandName ? serializeXml(islandName) : null);
		return xhr.responseText || "";
	}

	function loadXml(url, islandName) {
		var xhr = new XMLHttpRequest();
		var doc = xmlDocument(islandName);
		xhr.open("GET", url, false);
		xhr.send(null);
		if (trim(xhr.responseText) !== "" && doc && typeof doc.loadXML === "function") {
			doc.loadXML(xhr.responseText);
			return true;
		}
		return false;
	}

	function setDoneReturn() {
		var root = xmlRoot("RetData");
		if (root) {
			root.setAttribute("Done", "Y");
		}
		if (window.ITMSModernCompat) {
			window.ITMSModernCompat.returnModalValue(root);
		} else {
			window.returnValue = root;
			window.returnvalue = root;
		}
	}

	function table() {
		return document.getElementById("tblData");
	}

	function clearTable() {
		var grid = table();
		while (grid && grid.rows.length > 1) {
			grid.deleteRow(1);
		}
	}

	function appendCell(row, text, className, align) {
		var cell = row.insertCell();
		cell.textContent = text == null ? "" : String(text);
		cell.className = className || "";
		if (align) {
			cell.align = align;
		}
		return cell;
	}

	function appendInputCell(row, input) {
		var cell = row.insertCell();
		cell.className = "ExcelInputCell";
		cell.appendChild(input);
		return cell;
	}

	function createTextInput(name, data, readOnly, disabled, onblur) {
		var input = document.createElement("input");
		input.type = "text";
		input.name = name;
		input.size = name.indexOf("txtOrder") === 0 ? 11 : 40;
		input.maxLength = name.indexOf("txtOrder") === 0 ? 4 : 50;
		input.className = "FormElem";
		input.value = data || "";
		input.readOnly = !!readOnly;
		input.disabled = !!disabled;
		if (onblur) {
			input.onblur = onblur;
		}
		return input;
	}

	function createRow(index, node, isNewRow) {
		var grid = table();
		var row = grid.insertRow(grid.rows.length);
		var cell;
		var checkbox;
		var practiceDisabled = !isNewRow;
		var orderDisabled = !isNewRow;
		appendCell(row, index, "ExcelSerial", "center").width = 10;
		cell = appendCell(row, "", "ExcelDisplayCell", "center");
		checkbox = document.createElement("input");
		checkbox.type = "checkbox";
		checkbox.name = "ChkboxEdit" + index;
		checkbox.className = "Formelem";
		checkbox.value = "1";
		checkbox.disabled = isNewRow || value("hType") !== "PROCESS";
		cell.appendChild(checkbox);
		appendInputCell(row, createTextInput("txtPractice" + index, getAttr(node, "PROCESSNAME"), practiceDisabled, practiceDisabled, null));
		cell = appendInputCell(row, createTextInput("txtOrder" + index, getAttr(node, "ORDERNUMBER"), true, orderDisabled, function () {
			rowAdd(index);
		}));
		cell.align = "center";
	}

	function addRow() {
		var grid = table();
		var index = grid ? grid.rows.length : 1;
		var root = xmlRoot("OutData");
		var doc = xmlDocument("OutData");
		var node = doc.createElement("DETAILS");
		setAttr(node, "CTR", index);
		setAttr(node, "PROCESSCODE", "");
		setAttr(node, "PROCESSNAME", "");
		setAttr(node, "ORDERNUMBER", "");
		root.appendChild(node);
		createRow(index, node, true);
		setValue("hItemRows", index);
		setValue("hNoOfRows", index);
	}

	function rowAdd(index) {
		if (trim(value("txtPractice" + index)) === "") {
			return false;
		}
		if (trim(value("txtOrder" + index)) === "") {
			return false;
		}
		if (isNaN(Number(value("txtOrder" + index)))) {
			alert("Enter Numerals Only");
			return false;
		}
		addRow();
		return true;
	}

	function applyPracticeMode() {
		if (value("hType") !== "PRACTICE") {
			return;
		}
		detailNodes().forEach(function (node) {
			var index;
			if (getAttr(node, "PROCESSCODE") !== value("hPracticeCode")) {
				return;
			}
			index = getAttr(node, "CTR");
			if (field("txtPractice" + index)) {
				field("txtPractice" + index).disabled = false;
				field("txtPractice" + index).readOnly = false;
			}
			if (field("ChkboxEdit" + index)) {
				field("ChkboxEdit" + index).disabled = false;
			}
		});
	}

	function displayTable() {
		var index = 0;
		clearTable();
		detailNodes().forEach(function (node) {
			index += 1;
			createRow(index, node, false);
			setValue("hItemRows", index);
		});
		applyPracticeMode();
	}

	function showData(callFrom, processCode, practiceCode) {
		var passData = callFrom + ":" + processCode + ":" + practiceCode;
		upgradeXml();
		if (value("hCheck") !== "NEW") {
			if (!loadXml("XMLSelect.asp?sWho=ACTONLOAD&sPassData=" + encodeURIComponent(passData), "OutData")) {
				return false;
			}
			displayTable();
		}
		if (value("hType") === "ADDPRACTICE") {
			addRow();
		}
		return true;
	}

	function goToAction() {
		var checked = !!(field("ChkboxEdit") && field("ChkboxEdit").checked);
		var count = parseInt(value("hItemRows"), 10) || 0;
		if (value("hCheck") === "NEW") {
			return false;
		}
		for (var i = 1; i <= count; i += 1) {
			if (field("txtPractice" + i)) {
				field("txtPractice" + i).readOnly = !checked;
				if (value("hType") === "PROCESS" || value("hType") === "ADDPRACTICE") {
					field("txtPractice" + i).disabled = !checked;
				}
			}
			if (field("txtOrder" + i)) {
				field("txtOrder" + i).readOnly = !checked;
				field("txtOrder" + i).disabled = !checked;
			}
		}
		return false;
	}

	function deleteData(practiceCode, processCodes) {
		var response = postXml("AppActivityDelete.asp?sPassData=" + encodeURIComponent(practiceCode + ":" + processCodes), "");
		if (trim(response) !== "") {
			alert(response);
		} else {
			alert("Record Deleted Successfully");
		}
	}

	function deleteItem() {
		var selectedCodes = [];
		var count = parseInt(value("hItemRows"), 10) || 0;
		var root = xmlRoot("OutData");
		for (var i = 1; i <= count; i += 1) {
			if (field("ChkboxEdit" + i) && field("ChkboxEdit" + i).checked) {
				detailNodes().forEach(function (node) {
					if (parseInt(getAttr(node, "CTR"), 10) === i) {
						selectedCodes.push(getAttr(node, "PROCESSCODE"));
						root.removeChild(node);
					}
				});
			}
		}
		selectedCodes = selectedCodes.filter(function (code) {
			return trim(code) !== "";
		});
		if (!selectedCodes.length) {
			alert("Select any One option for Delete");
			return false;
		}
		if (confirm("This will delete the Activity which u selected. Do you want to continue?")) {
			deleteData(value("hProcessCode"), selectedCodes.join(","));
		}
		displayTable();
		return false;
	}

	function detailByCtr(index) {
		var found = null;
		detailNodes().forEach(function (node) {
			if (parseInt(getAttr(node, "CTR"), 10) === index) {
				found = node;
			}
		});
		return found;
	}

	function maxOrder() {
		var max = 0;
		detailNodes().forEach(function (node) {
			var current = parseInt(getAttr(node, "ORDERNUMBER"), 10);
			if (!isNaN(current) && current > max) {
				max = current;
			}
		});
		return max;
	}

	function rowsToValidate(callFrom) {
		var rows = parseInt(value("hItemRows"), 10) || 0;
		var addRows;
		if (value("hCheck") === "NEW") {
			rows = 1;
			setValue("hItemRows", "1");
		}
		if (callFrom === "ADDPRACTICE") {
			addRows = parseInt(value("hNoOfRows"), 10) || rows;
			rows = addRows !== 1 ? addRows - 1 : addRows;
		}
		return rows;
	}

	function validateAndUpdateRows(callFrom) {
		var rows = rowsToValidate(callFrom);
		var previousOrder = "0";
		var root = xmlRoot("OutData");
		maxOrder();
		for (var i = 1; i <= rows; i += 1) {
			var node = detailByCtr(i);
			var practice = trim(value("txtPractice" + i));
			var order = trim(value("txtOrder" + i));
			if (!node) {
				continue;
			}
			if (practice === "") {
				alert("Enter Practice Name");
				return false;
			}
			if (order === "") {
				alert("Enter Order No");
				return false;
			}
			if (Number(order) <= 0) {
				alert("Order Number Must be Greater Than Zero");
				return false;
			}
			if (isNaN(Number(order))) {
				alert("Enter Numerals Only");
				return false;
			}
			if (previousOrder === order) {
				alert("Same Order Number is Not allowed");
				return false;
			}
			previousOrder = order;
			setAttr(node, "ORDERNUMBER", order);
			if (trim(callFrom) === "PRACTICE" || trim(callFrom) === "ADDPRACTICE") {
				setAttr(node, "PROCESSNAME", practice);
			}
		}
		if (callFrom === "ADDPRACTICE") {
			detailNodes().forEach(function (node) {
				if (parseInt(getAttr(node, "CTR"), 10) > rows && trim(getAttr(node, "PROCESSNAME")) === "" && trim(getAttr(node, "ORDERNUMBER")) === "") {
					root.removeChild(node);
				}
			});
		}
		return true;
	}

	function saveData(callFrom) {
		var passData;
		var response;
		var processCode = value("hPracticeCode");
		var practiceCode = value("hProcessCode");
		var processName = value("txtPrcessName");
		if (!validateAndUpdateRows(callFrom)) {
			return false;
		}
		passData = callFrom + ":" + processCode + ":" + practiceCode + ":" + processName;
		response = postXml("AppActivityUpdate.asp?sPassData=" + encodeURIComponent(passData), "OutData");
		if (trim(response) !== "") {
			alert(response);
			return false;
		}
		alert("Record Udated Successfully");
		setDoneReturn();
		window.close();
		return true;
	}

	function closePopup() {
		window.close();
	}

	function checkSubmit() {
		if (detailNodes().length > 0) {
			setDoneReturn();
		}
		window.close();
	}

	function install() {
		window.GoToAction = goToAction;
		window.addRow = addRow;
		window.RowAdd = rowAdd;
		window.ShowData = showData;
		window.DisplayTable = displayTable;
		window.ClearTable = clearTable;
		window.DeleteItem = deleteItem;
		window.Max = maxOrder;
		window.SaveData = saveData;
		window.DeleteData = deleteData;
		window.CheckSubmit = checkSubmit;
		window.ClosePopup = closePopup;
	}

	window.ITMSAppActivityDetailsPopupCompat = {
		install: install
	};
}(window, document));
