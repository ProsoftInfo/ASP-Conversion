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

	function postXml(url, islandName) {
		var xhr = new XMLHttpRequest();
		xhr.open("POST", url, false);
		xhr.setRequestHeader("Content-Type", "text/xml");
		xhr.send(islandName ? serializeXml(islandName) : null);
		return xhr.responseText || "";
	}

	function parseXml(text) {
		return new DOMParser().parseFromString(text || "<Root/>", "text/xml").documentElement;
	}

	function requestRoot(url) {
		var xhr = new XMLHttpRequest();
		xhr.open("GET", url, false);
		xhr.send(null);
		return trim(xhr.responseText) === "" ? null : (xhr.responseXML && xhr.responseXML.documentElement || parseXml(xhr.responseText));
	}

	function resetSelect(selectName) {
		var select = field(selectName);
		if (select) {
			select.options.length = 1;
			select.selectedIndex = 0;
		}
	}

	function fillSelect(selectName, root, textIndex, valueIndex) {
		var select = field(selectName);
		if (!select || !root) {
			return false;
		}
		select.options.length = 1;
		childElements(root).forEach(function (node) {
			var attrs = node.attributes;
			select.options[select.options.length] = new Option(attrs.item(textIndex).value, attrs.item(valueIndex).value);
		});
		select.selectedIndex = 0;
		return true;
	}

	function selectedOptionText(selectName) {
		var select = field(selectName);
		return select && select.options && select.selectedIndex >= 0 ? select.options[select.selectedIndex].text : "";
	}

	function selectedActivityOptions() {
		var select = field("selActivtyDes");
		var result = [];
		if (!select) {
			return result;
		}
		Array.prototype.forEach.call(select.options, function (option) {
			if (option.selected && option.value !== "S") {
				result.push(option);
			}
		});
		return result;
	}

	function mappingNodes() {
		return childElements(xmlRoot("OutData")).filter(function (node) {
			return node.nodeName === "ACTIVITYMAPPING";
		});
	}

	function activityNodes(mapping) {
		return childElements(mapping).filter(function (node) {
			return node.nodeName === "ACTIVITY";
		});
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

	function appendHeader(text) {
		var row = table().insertRow(table().rows.length);
		var cell = appendCell(row, String(text || "").toUpperCase(), "ExcelHeaderCell", "left");
		cell.colSpan = 3;
	}

	function displayTable() {
		var mappingIndex = 0;
		var serial = 1;
		var previousProcess = "";
		clearTable();
		mappingNodes().forEach(function (mapping) {
			mappingIndex += 1;
			setAttr(mapping, "CTR", mappingIndex);
			if (getAttr(mapping, "APPNAME") !== previousProcess) {
				previousProcess = getAttr(mapping, "APPNAME");
				appendHeader(previousProcess);
			}
			appendHeader(getAttr(mapping, "PROCESSNAME"));
			activityNodes(mapping).forEach(function (activity) {
				var row = table().insertRow(table().rows.length);
				var cell;
				var checkbox;
				setAttr(activity, "ACTCTR", serial);
				appendCell(row, serial, "ExcelSerial", "center");
				cell = appendCell(row, "", "ExcelDisplayCell", "center");
				checkbox = document.createElement("input");
				checkbox.type = "checkbox";
				checkbox.name = "mChkValue" + serial;
				checkbox.className = "Formelem";
				cell.appendChild(checkbox);
				appendCell(row, getAttr(activity, "NAME") + "-" + getAttr(activity, "TNAME"), "ExcelDisplayCell", "left");
				serial += 1;
			});
		});
		setValue("hItemRows", serial);
	}

	function showData(callFrom, data) {
		var passData = callFrom + ":" + value("hRoleID") + ":" + data;
		if (callFrom === "FROMPRACTICE") {
			passData = callFrom + ":" + value("hRoleID") + ":" + value("selProcess") + ":" + data;
		}
		if (loadXml("XMLSelect.asp?sWho=PPAONLOAD&sPassData=" + encodeURIComponent(passData), "OutData")) {
			displayTable();
		}
	}

	function populatePractice(source) {
		var selectedValue = source && source.value;
		var root;
		showData("FROMPROCESS", selectedValue);
		if (!source || source.selectedIndex === 0) {
			resetSelect("selPractice");
			resetSelect("selActivtyDes");
			return false;
		}
		root = requestRoot("XMLSelect.asp?sWho=PR&sProcess=" + encodeURIComponent(selectedValue));
		fillSelect("selPractice", root, 1, 0);
		resetSelect("selActivtyDes");
		return true;
	}

	function populateActivities(source) {
		var practiceValue = source && source.value;
		var processValue = value("selProcess");
		var root;
		if (!source || source.selectedIndex === 0) {
			resetSelect("selActivtyDes");
			return false;
		}
		root = requestRoot("XMLSelect.asp?sWho=PPA&sProcess=" + encodeURIComponent(practiceValue + ":" + processValue + ":" + value("hRoleID")));
		fillSelect("selActivtyDes", root, 2, 1);
		showData("FROMPRACTICE", practiceValue);
		setValue("hLastSelectedPractice", practiceValue);
		setValue("hLastSelectedProcess", processValue);
		return true;
	}

	function firstMapping(appCode, appName, processCode, processName) {
		var root = xmlRoot("OutData");
		var doc = xmlDocument("OutData");
		var mapping = mappingNodes()[0];
		if (!mapping) {
			mapping = doc.createElement("ACTIVITYMAPPING");
			setAttr(mapping, "CTR", "0");
			setAttr(mapping, "APPCODE", trim(appCode));
			setAttr(mapping, "APPNAME", appName);
			setAttr(mapping, "PROCESSCODE", trim(processCode));
			setAttr(mapping, "PROCESSNAME", processName);
			root.appendChild(mapping);
		}
		return mapping;
	}

	function duplicateByName(mapping, activityName, templateName, includeTemplate) {
		return activityNodes(mapping).some(function (activity) {
			if (includeTemplate) {
				return trim(getAttr(activity, "NAME")) + "-" + trim(getAttr(activity, "TNAME")) === trim(activityName) + "-" + trim(templateName);
			}
			return trim(getAttr(activity, "NAME")) === trim(activityName);
		});
	}

	function appendActivity(mapping, activityCode, activityName, tempNo, tempName) {
		var node = xmlDocument("OutData").createElement("ACTIVITY");
		setAttr(node, "ACTCTR", "0");
		setAttr(node, "CODE", activityCode);
		setAttr(node, "NAME", activityName);
		setAttr(node, "TCODE", tempNo);
		setAttr(node, "TNAME", tempName);
		mapping.appendChild(node);
	}

	function firstTemplateFallback(activityDataRoot, activityName) {
		var result = { tempNo: "1", tempName: activityName };
		childElements(activityDataRoot).some(function (activity) {
			return childElements(activity).some(function (template) {
				if (template.nodeName === "Template") {
					result.tempNo = getAttr(template, "TempNo");
					result.tempName = getAttr(template, "TempName");
					return true;
				}
				return false;
			});
		});
		return result;
	}

	function selectedActivityCodesForTemplateRequest(options) {
		return options.map(function (option) {
			return String(option.value).split(":")[0];
		}).join(",");
	}

	function processTemplateSelection(mapping, roleRoot) {
		if (!roleRoot || !roleRoot.hasChildNodes()) {
			return false;
		}
		childElements(roleRoot).forEach(function (role) {
			var activityName;
			var tempName;
			if (role.nodeName !== "Role") {
				return;
			}
			activityName = getAttr(role, "ActName");
			tempName = getAttr(role, "TempName");
			if (duplicateByName(mapping, activityName, tempName, true)) {
				alert("Already Selected Activity is Mapped");
				return;
			}
			appendActivity(mapping, getAttr(role, "ActCode"), activityName, getAttr(role, "TempNo"), tempName);
		});
		displayTable();
		saveData();
		return true;
	}

	function openTemplatePopup(appCode, processCode, activityCodes, callback) {
		var url = "ActRolTempSelPop.asp?AppCode=" + encodeURIComponent(appCode) + "&ProcessCode=" + encodeURIComponent(processCode) + "&ActCode=" + encodeURIComponent(activityCodes);
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			return window.ITMSModernCompat.openModalDialog(url, "", "dialogWidth:500px;dialogHeight:500px", callback);
		}
		return window.open(url, "_blank", "width=500,height=500,resizable=yes,scrollbars=yes");
	}

	function addEntry() {
		var appCode = value("selProcess");
		var processCode = value("selPractice");
		var options = selectedActivityOptions();
		var appName;
		var processName;
		var mapping;
		var activityCodes;
		var activityRoot;
		var eligible;
		if (appCode === "S") {
			alert("Select Process");
			field("selProcess").focus();
			return false;
		}
		if (processCode === "S") {
			alert("Select Practice");
			field("selPractice").focus();
			return false;
		}
		if (!options.length) {
			alert("Select Activity");
			field("selActivtyDes").focus();
			return false;
		}
		appName = selectedOptionText("selProcess");
		processName = selectedOptionText("selPractice");
		mapping = firstMapping(appCode, appName, processCode, processName);
		activityCodes = selectedActivityCodesForTemplateRequest(options);
		loadXml("XMLGetActTemp.asp?AppCode=" + encodeURIComponent(appCode) + "&ProcessCode=" + encodeURIComponent(processCode) + "&ActCode=" + encodeURIComponent(activityCodes), "ActivityData");
		activityRoot = xmlRoot("ActivityData");
		eligible = getAttr(activityRoot, "Eligible");
		if (eligible === "Y") {
			openTemplatePopup(appCode, processCode, activityCodes, function (roleRoot) {
				processTemplateSelection(mapping, roleRoot);
			});
			return false;
		}
		options.forEach(function (option) {
			var activityCode = option.value;
			var activityName = option.text;
			var fallback = firstTemplateFallback(activityRoot, activityName);
			if (duplicateByName(mapping, activityName, fallback.tempName, false)) {
				alert("Already Selected Activity is Mapped");
				return;
			}
			appendActivity(mapping, activityCode, activityName, fallback.tempNo, fallback.tempName);
		});
		displayTable();
		saveData();
		return false;
	}

	function restoreSelections() {
		var process = field("selProcess");
		var practice = field("selPractice");
		if (practice) {
			Array.prototype.forEach.call(practice.options, function (option, index) {
				if (option.value === value("hLastSelectedPractice")) {
					practice.selectedIndex = index;
				}
			});
		}
		if (process) {
			Array.prototype.forEach.call(process.options, function (option, index) {
				if (option.value === value("hLastSelectedProcess")) {
					process.selectedIndex = index;
				}
			});
		}
	}

	function saveData() {
		var response = postXml("RoleActivityMappingInsert.asp?sPassData=" + encodeURIComponent("ADD:" + value("hRoleID")), "OutData");
		if (trim(response) !== "") {
			alert(response);
		} else {
			alert("Record Added Successfully");
		}
		restoreSelections();
		populateActivities(field("selPractice"));
		clearTable();
		displayTable();
	}

	function deleteData(applicationCode, processCode, activityCode, templateNo) {
		var passData = "DEL:" + processCode + ":" + activityCode + ":" + applicationCode + ":" + value("hRoleID") + ":" + templateNo;
		var response = postXml("RoleActivityMappingInsert.asp?sPassData=" + encodeURIComponent(passData), "OutData");
		if (trim(response) !== "") {
			alert(response);
		} else {
			alert("Record Deleted Successfully");
		}
		restoreSelections();
		populateActivities(field("selPractice"));
		clearTable();
		displayTable();
	}

	function deleteItem() {
		var rows = (parseInt(value("hItemRows"), 10) || 1) - 1;
		var applicationCode = "";
		var processCode = "";
		var activityCodes = [];
		var templateNos = [];
		for (var i = 1; i <= rows; i += 1) {
			if (!field("mChkValue" + i) || !field("mChkValue" + i).checked) {
				continue;
			}
			mappingNodes().forEach(function (mapping) {
				activityNodes(mapping).forEach(function (activity) {
					if (parseInt(getAttr(activity, "ACTCTR"), 10) === i) {
						applicationCode = getAttr(mapping, "APPCODE");
						processCode = getAttr(mapping, "PROCESSCODE");
						activityCodes.push(getAttr(activity, "CODE"));
						templateNos.push(getAttr(activity, "TCODE"));
					}
				});
			});
		}
		if (!activityCodes.length) {
			alert("Select any One option for Delete");
			return false;
		}
		if (confirm("This will delete the Activity which u selected. Do you want to continue?")) {
			deleteData(applicationCode, processCode, activityCodes.join(","), templateNos.join(","));
		}
		return false;
	}

	function checkSubmit() {
		var retRoot = xmlRoot("RetData");
		if (mappingNodes().length > 0 && retRoot) {
			retRoot.setAttribute("Done", "Y");
			if (window.ITMSModernCompat) {
				window.ITMSModernCompat.returnModalValue(retRoot);
			} else {
				window.returnValue = retRoot;
				window.returnvalue = retRoot;
			}
		}
		window.close();
	}

	function install() {
		upgradeXml();
		window.PopulatePractice = populatePractice;
		window.PopulateActivities = populateActivities;
		window.AddEntry = addEntry;
		window.DisplayTable = displayTable;
		window.ClearTable = clearTable;
		window.DeleteItem = deleteItem;
		window.ShowData = showData;
		window.SaveData = saveData;
		window.DeleteData = deleteData;
		window.CheckSubmit = checkSubmit;
	}

	window.ITMSAppRoleActivityMappingPopupCompat = {
		install: install
	};
}(window, document));
