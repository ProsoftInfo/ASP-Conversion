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

	function getAttr(node, name) {
		return node && node.getAttribute(name) || "";
	}

	function clearTable() {
		var table = document.getElementById("tblTempAct");
		while (table && table.rows.length > 2) {
			table.deleteRow(2);
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

	function displayTable() {
		var table = document.getElementById("tblTempAct");
		var serial = 0;
		if (!table) {
			return false;
		}
		clearTable();
		childElements(xmlRoot("ActivityData")).forEach(function (activity) {
			var actCode = getAttr(activity, "ActCode");
			var actName = getAttr(activity, "ActName");
			var tempCount = parseInt(getAttr(activity, "TempCnt"), 10) || 0;
			var row;
			var cell;
			var checkbox;
			if (activity.nodeName !== "Activity" || tempCount <= 1) {
				return;
			}
			serial += 1;
			row = table.insertRow(table.rows.length);
			appendCell(row, serial, "ExcelSerial", "center");
			appendCell(row, "", "ExcelDisplayCell", "center");
			appendCell(row, actName, "ExcelDisplayCell");
			childElements(activity).forEach(function (template) {
				var tempNo;
				if (template.nodeName !== "Template") {
					return;
				}
				tempNo = getAttr(template, "TempNo");
				row = table.insertRow(table.rows.length);
				appendCell(row, "", "ExcelSerial", "center");
				cell = appendCell(row, "", "ExcelDisplayCell", "center");
				checkbox = document.createElement("input");
				checkbox.type = "checkbox";
				checkbox.name = "ChkTempZ" + actCode + "Z" + tempNo;
				checkbox.checked = trim(getAttr(template, "Select")) === "Y";
				cell.appendChild(checkbox);
				appendCell(row, "   " + getAttr(template, "TempName"), "ExcelDisplayCell");
			});
		});
		return true;
	}

	function loadActivities() {
		var xhr = new XMLHttpRequest();
		var doc = xmlDocument("ActivityData");
		var url = "XMLGetActTemp.asp?AppCode=" + encodeURIComponent(field("hAppCode").value) + "&ProcessCode=" + encodeURIComponent(field("hProcessCode").value) + "&ActCode=" + encodeURIComponent(field("hActCode").value);
		xhr.open("GET", url, false);
		xhr.send(null);
		if (trim(xhr.responseText) !== "" && doc && typeof doc.loadXML === "function") {
			doc.loadXML(xhr.responseText);
		}
	}

	function init() {
		upgradeXml();
		loadActivities();
		clearTable();
		displayTable();
	}

	function appendRole(roleRoot, actCode, actName, tempNo, tempName) {
		var role = xmlDocument("RoleData").createElement("Role");
		role.setAttribute("ActCode", actCode);
		role.setAttribute("ActName", actName);
		role.setAttribute("TempNo", tempNo);
		role.setAttribute("TempName", tempName);
		roleRoot.appendChild(role);
	}

	function checkSubmit() {
		var roleRoot = xmlRoot("RoleData");
		while (roleRoot && roleRoot.firstChild) {
			roleRoot.removeChild(roleRoot.firstChild);
		}
		childElements(xmlRoot("ActivityData")).forEach(function (activity) {
			var actCode = getAttr(activity, "ActCode");
			var actName = getAttr(activity, "ActName");
			var tempCount = parseInt(getAttr(activity, "TempCnt"), 10) || 0;
			if (activity.nodeName !== "Activity") {
				return;
			}
			if (tempCount > 1) {
				childElements(activity).forEach(function (template) {
					var tempNo = getAttr(template, "TempNo");
					var checkbox = field("ChkTempZ" + actCode + "Z" + tempNo);
					if (template.nodeName === "Template" && checkbox && checkbox.checked && trim(getAttr(template, "Select")) === "N") {
						appendRole(roleRoot, actCode, actName, tempNo, getAttr(template, "TempName"));
					}
				});
			} else {
				appendRole(roleRoot, actCode, actName, "1", actName);
			}
		});
		if (roleRoot) {
			roleRoot.setAttribute("Done", "Y");
		}
		if (window.ITMSModernCompat) {
			window.ITMSModernCompat.returnModalValue(roleRoot);
		} else {
			window.returnValue = roleRoot;
			window.returnvalue = roleRoot;
		}
		window.close();
		return false;
	}

	function install() {
		window.Init = init;
		window.DisplayTable = displayTable;
		window.ClearTable = clearTable;
		window.CheckSubmit = checkSubmit;
	}

	window.ITMSActRolTempSelPopCompat = {
		install: install
	};
}(window, document));
