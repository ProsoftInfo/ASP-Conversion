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

	function submitForm() {
		var frm = form();
		if (frm) {
			frm.submit();
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

	function clearChildren(node) {
		while (node && node.firstChild) {
			node.removeChild(node.firstChild);
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

	function requestRoot(url) {
		var xhr = new XMLHttpRequest();
		var doc;
		xhr.open("GET", url, false);
		xhr.send(null);
		if (trim(xhr.responseText) === "") {
			return null;
		}
		doc = xhr.responseXML && xhr.responseXML.documentElement ? xhr.responseXML : new DOMParser().parseFromString(xhr.responseText, "text/xml");
		return doc.documentElement;
	}

	function loadXmlIsland(name, url) {
		var xhr = new XMLHttpRequest();
		var data;
		var doc;
		var xmlText;
		xhr.open("GET", url, false);
		xhr.send(null);
		xmlText = xhr.responseText || "";
		if (trim(xmlText) === "" && xhr.responseXML && xhr.responseXML.documentElement) {
			xmlText = new XMLSerializer().serializeToString(xhr.responseXML);
		}
		if (trim(xmlText) === "") {
			return null;
		}
		data = xmlObject(name);
		if (data && typeof data.loadXML === "function") {
			data.loadXML(xmlText);
			return xmlRoot(name);
		}
		doc = xhr.responseXML && xhr.responseXML.documentElement ? xhr.responseXML : new DOMParser().parseFromString(xmlText, "text/xml");
		if (data) {
			data._doc = doc;
		}
		return doc.documentElement;
	}

	function table() {
		return document.getElementById("tblData");
	}

	function clearTableRows() {
		var grid = table();
		while (grid && grid.rows.length > 1) {
			grid.deleteRow(1);
		}
	}

	function addCell(row, className, align, text) {
		var cell = row.insertCell();
		cell.className = className || "";
		if (align) {
			cell.align = align;
		}
		cell.innerHTML = text == null ? "" : String(text);
		return cell;
	}

	function addLegacyHeader(text) {
		var grid = table();
		var row = grid && grid.insertRow(grid.rows.length);
		var cell;
		if (!row) {
			return;
		}
		cell = addCell(row, "ExcelHeaderCell", "left", String(text || "").toUpperCase());
		cell.colSpan = 3;
	}

	function addLegacyActivityRow(index, rowValue, activityName) {
		var grid = table();
		var row = grid && grid.insertRow(grid.rows.length);
		if (!row) {
			return;
		}
		addCell(row, "ExcelSerial", "center", index);
		addCell(row, "ExcelDisplayCell", "center", '<input type="checkbox" name="mChkValue' + index + '" value="' + rowValue + '" size="11" class="Formelem">');
		addCell(row, "ExcelDisplayCell", "left", activityName);
	}

	function legacyTomapRows(root) {
		var rows = [];
		childElements(root).forEach(function (node) {
			var roleId;
			var roleName;
			if (node.nodeName === "ROLE") {
				roleId = node.getAttribute("ROLEID") || value("hRoleID") || value("selRole") || "0";
				roleName = node.getAttribute("ROLENAME") || "";
				rows.push({
					header: roleName
				});
				childElements(node).forEach(function (child) {
					if (child.nodeName === "TOMAP") {
						rows.push({
							node: child,
							roleId: roleId
						});
					}
				});
			} else if (node.nodeName === "TOMAP") {
				rows.push({
					node: node,
					roleId: value("hRoleID") || value("selRole") || "0"
				});
			}
		});
		return rows;
	}

	function legacyRowValue(roleId, node) {
		return [
			roleId || "0",
			node.getAttribute("APPCode") || "",
			node.getAttribute("PRCode") || "",
			node.getAttribute("ACCode") || "",
			node.getAttribute("TempNo") || node.getAttribute("ActivityTemplateNo") || "0"
		].join(":");
	}

	function displayLegacyTable() {
		var rows = legacyTomapRows(xmlRoot("OutData"));
		var serial = 1;
		var processName = "";
		var practiceName = "";
		clearTableRows();
		rows.forEach(function (entry) {
			var node;
			if (entry.header !== undefined) {
				addLegacyHeader(entry.header);
				processName = "";
				practiceName = "";
				return;
			}
			node = entry.node;
			if ((node.getAttribute("APPName") || "") !== processName) {
				processName = node.getAttribute("APPName") || "";
				addLegacyHeader(processName);
			}
			if ((node.getAttribute("PAName") || "") !== practiceName) {
				practiceName = node.getAttribute("PAName") || "";
				addLegacyHeader(practiceName);
			}
			addLegacyActivityRow(serial, legacyRowValue(entry.roleId, node), node.getAttribute("ACName") || "");
			serial += 1;
		});
		setValue("hItemRows", serial - 1);
	}

	function fillPracticeSelect(processCode) {
		var select = field("selActivityType");
		var root;
		if (processCode === "0" || processCode === "S") {
			alert("Select Prcoess Name");
			if (field("selProcess")) {
				field("selProcess").focus();
			}
			return false;
		}
		if (!select) {
			return false;
		}
		select.options.length = 1;
		root = requestRoot("XMLSelect.asp?sWho=PR&sProcess=" + encodeURIComponent(processCode));
		if (!root) {
			alert("No Practice defined for the Process Selected");
			if (field("selProcess")) {
				field("selProcess").focus();
			}
			return false;
		}
		childElements(root).forEach(function (node) {
			select.options[select.options.length] = new Option(node.attributes.item(1).value, node.attributes.item(0).value);
		});
		return true;
	}

	function selectedPracticeCodes() {
		var select = field("selActivityType");
		var values = [];
		if (!select) {
			return "";
		}
		Array.prototype.forEach.call(select.options, function (option) {
			if (option.selected && option.value !== "S") {
				values.push(option.value);
			}
		});
		return values.join(",");
	}

	function validateFilters(hasRole) {
		if (hasRole) {
			setValue("hRoleID", value("selRole"));
		}
		setValue("hProcessCode", value("selProcess"));
		setValue("hPracticeCode", selectedPracticeCodes());
		submitForm();
	}

	function paginate(nPage, hasRole) {
		setValue("hPageSelection", nPage);
		if (hasRole) {
			setValue("selRole", value("hRoleID"));
		}
		setValue("selProcess", value("hProcessCode"));
		if (field("selActivityType")) {
			field("selActivityType").value = value("hPracticeCode");
		}
		submitForm();
	}

	function resetFilters() {
		setValue("hProcessCode", "");
		setValue("hPracticeCode", "");
		if (field("selProcess")) {
			field("selProcess").selectedIndex = 0;
		}
		if (field("selActivityType")) {
			field("selActivityType").options.length = 1;
		}
		submitForm();
	}

	function appendLegacyActivity(root, rowValue, fallbackRoleId) {
		var parts = String(rowValue || "").split(":");
		var node = xmlDocument("SelectedData").createElement("ACTIVITY");
		node.setAttribute("RoleID", parts[0] || fallbackRoleId || "0");
		node.setAttribute("APPCode", parts[1] || "");
		node.setAttribute("PRCode", parts[2] || "");
		node.setAttribute("ACCode", parts[3] || "");
		node.setAttribute("TempNo", parts[4] || "0");
		root.appendChild(node);
	}

	function legacyLoadMapped(userId, roleId) {
		loadXmlIsland("OutData", "XMLSelect.asp?sWho=UAMAP&sPassData=" + encodeURIComponent(userId + ":" + roleId));
		displayLegacyTable();
	}

	function legacyDeleteItem() {
		var root = xmlRoot("SelectedData");
		var selected = selectedRows("mChkValue", "hItemRows");
		var response;
		clearChildren(root);
		root.setAttribute("RoleID", value("selRole"));
		root.setAttribute("UserID", value("hEmpID"));
		selected.forEach(function (rowValue) {
			appendLegacyActivity(root, rowValue, value("selRole"));
		});
		if (!selectedActivityCount(root)) {
			alert("Select Activity For deletion");
			return false;
		}
		response = postXml("RoleActivityMappingInsert.asp?sPassData=DELAPPUSERROLE", "SelectedData");
		if (trim(response) !== "") {
			alert(response);
		} else {
			alert("Record Deleted Successfully");
			legacyLoadMapped(value("hEmpID"), value("selRole"));
		}
		return false;
	}

	function legacyShowActivityDetails() {
		var passData;
		if (value("selRole") === "S") {
			alert("Select Role");
			if (field("selRole")) {
				field("selRole").focus();
			}
			return false;
		}
		if (value("selIType") === "S") {
			alert("Select Item Type");
			if (field("selIType")) {
				field("selIType").focus();
			}
			return false;
		}
		passData = value("hEmpID") + ":" + value("selRole") + ":" + value("selIType");
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog("Copy%20of%20AppActivityRolePopUp.asp?PassData=" + encodeURIComponent(passData), "", "dialogHeight:320px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No;scroll:no", function (returnedValue) {
				var returnedRoot = getReturnedRoot(returnedValue);
				if (returnedRoot && returnedRoot.getAttribute("Done") === "Y") {
					legacyLoadMapped(value("hEmpID"), value("selRole"));
				}
			});
		}
		return false;
	}

	function legacyPopupShowData() {
		var passData = value("hEmpID") + ":" + value("hRoleID") + ":S:S";
		loadXmlIsland("OutData", "XMLSelect.asp?sWho=UANM&sPassData=" + encodeURIComponent(passData));
		displayLegacyTable();
	}

	function legacyPopupCheckSubmit() {
		var root = xmlRoot("SelectedData");
		var selected = selectedRows("mChkValue", "hItemRows");
		var response;
		clearChildren(root);
		root.setAttribute("RoleID", value("hRoleID"));
		root.setAttribute("UserID", value("hEmpID"));
		root.setAttribute("ItemType", value("hItemType"));
		selected.forEach(function (rowValue) {
			appendLegacyActivity(root, rowValue, value("hRoleID"));
		});
		if (!selectedActivityCount(root)) {
			window.close();
			return false;
		}
		response = postXml("RoleActivityMappingInsert.asp?sPassData=ADDAPPUSERROLE", "SelectedData");
		if (trim(response) !== "") {
			alert(response);
		} else {
			alert("Record Added Successfully");
			xmlRoot("RetData").setAttribute("Done", "Y");
			if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
				window.ITMSModernCompat.returnModalValue(xmlRoot("RetData"));
			}
		}
		window.close();
		return false;
	}

	function appendActivity(root, parts, mapping) {
		var node = xmlDocument("SelectedData").createElement("ACTIVITY");
		Object.keys(mapping).forEach(function (name) {
			node.setAttribute(name, parts[mapping[name]] || "");
		});
		root.appendChild(node);
	}

	function selectedRows(prefix, countName) {
		var rows = parseInt(value(countName), 10) || 0;
		var selected = [];
		for (var i = 1; i <= rows; i += 1) {
			var item = field(prefix + i);
			if (item && item.checked) {
				selected.push(item.value);
			}
		}
		return selected;
	}

	function selectedActivityCount(root) {
		return childElements(root).filter(function (node) {
			return node.nodeName === "ACTIVITY";
		}).length;
	}

	function deleteAssignedActivities() {
		var root = xmlRoot("SelectedData");
		var selected = selectedRows("Chkbox", "hCnt");
		var response;
		clearChildren(root);
		root.setAttribute("UserID", value("hEmpID"));
		selected.forEach(function (rowValue) {
			appendActivity(root, rowValue.split(":"), {
				RoleID: 0,
				APPCode: 1,
				PRCode: 2,
				ACCode: 3,
				TempNo: 4
			});
		});
		if (!selectedActivityCount(root)) {
			alert("Select Activity For deletion");
			return false;
		}
		response = postXml("RoleActivityMappingInsert.asp?sPassData=DELAPPUSERROLE", "SelectedData");
		if (trim(response) !== "") {
			alert(response);
		} else {
			alert("Record Deleted Successfully");
			submitForm();
		}
		return false;
	}

	function goBackToGrid() {
		var frm = form();
		frm.action = "ApplicationUserGrid.asp";
		submitForm();
	}

	function showActivityDetails() {
		var frm = form();
		frm.action = "AppActivityRolePopUp.asp?PassData=" + encodeURIComponent(value("hEmpID"));
		submitForm();
	}

	function collectPopupSelection() {
		var root = xmlRoot("SelectedData");
		var selected = selectedRows("mchkValue", "hCnt");
		clearChildren(root);
		root.setAttribute("UserID", value("hEmpID"));
		selected.forEach(function (rowValue) {
			appendActivity(root, rowValue.split(":"), {
				RoleID: 0,
				APPCode: 1,
				PRCode: 2,
				ACCode: 3,
				TempNo: 4
			});
		});
		return selectedActivityCount(root) > 0;
	}

	function getReturnedRoot(value) {
		return value && (value.documentElement || value.XMLDocument && value.XMLDocument.documentElement || value);
	}

	function savePopupData() {
		var root = xmlRoot("SelectedData");
		var open;
		if (!selectedActivityCount(root)) {
			alert("Select Any One Activity For Add");
			return false;
		}
		open = window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog;
		if (open) {
			window.ITMSModernCompat.openModalDialog("ItemTypeSelection.asp", "", "dialogHeight:250px;dialogWidth:240px;center:Yes;help:No;resizable:No;status:No;scroll:no", function (returnedValue) {
				var returnedRoot = getReturnedRoot(returnedValue);
				var response;
				if (!returnedRoot || returnedRoot.getAttribute("Done") === "") {
					alert("Select Item Type");
					return;
				}
				setValue("hItemTypeID", returnedRoot.getAttribute("ItemTypeID"));
				root.setAttribute("ItemType", value("hItemTypeID"));
				response = postXml("RoleActivityMappingInsert.asp?sPassData=ADDAPPUSERROLE", "SelectedData");
				if (trim(response) !== "") {
					alert(response);
				} else {
					alert("Record Added Successfully");
					xmlRoot("RetData").setAttribute("Done", "Y");
					submitForm();
				}
			});
		} else {
			alert("Popup support is not available.");
		}
		return false;
	}

	function addSelectedActivities() {
		if (collectPopupSelection()) {
			return savePopupData();
		}
		alert("Select Any One Activity For Add");
		return false;
	}

	function finalSubmit() {
		var frm = form();
		frm.action = "AppActivityRole.asp?EmpNo=" + encodeURIComponent(value("hEmpID"));
		submitForm();
	}

	function installMain() {
		upgradeXml();
		window.CheckSubmit = goBackToGrid;
		window.Validate = function () {
			validateFilters(false);
		};
		window.PopulatePractice = fillPracticeSelect;
		window.Paginate = function (nPage) {
			paginate(nPage, false);
		};
		window.DeleteItem = deleteAssignedActivities;
		window.ShowActivityDetails = showActivityDetails;
		window.ChkReset = resetFilters;
	}

	function installPopup() {
		upgradeXml();
		window.ShowData = function () {};
		window.DisplayTable = function () {};
		window.ClearTable = function () {};
		window.CheckSubmit = addSelectedActivities;
		window.FinalSubmit = finalSubmit;
		window.PopulatePractice = fillPracticeSelect;
		window.Validate = function () {
			validateFilters(true);
		};
		window.Paginate = function (nPage) {
			paginate(nPage, true);
		};
	}

	function installLegacyMain() {
		upgradeXml();
		window.populateList = legacyLoadMapped;
		window.DisplayTable = displayLegacyTable;
		window.ClearTable = clearTableRows;
		window.CheckSubmit = goBackToGrid;
		window.ShowActivityDetails = legacyShowActivityDetails;
		window.DeleteItem = legacyDeleteItem;
	}

	function installLegacyPopup() {
		upgradeXml();
		window.ShowData = legacyPopupShowData;
		window.DisplayTable = displayLegacyTable;
		window.ClearTable = clearTableRows;
		window.CheckSubmit = legacyPopupCheckSubmit;
		window.addEventListener("beforeunload", function () {
			if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
				window.ITMSModernCompat.returnModalValue(xmlRoot("RetData"));
			}
		});
	}

	window.ITMSAppActivityRoleCompat = {
		installMain: installMain,
		installPopup: installPopup,
		installLegacyMain: installLegacyMain,
		installLegacyPopup: installLegacyPopup
	};
}(window, document));
