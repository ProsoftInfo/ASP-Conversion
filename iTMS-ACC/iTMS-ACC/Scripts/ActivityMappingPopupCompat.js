(function (window, document) {
	"use strict";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.formname || document.forms.formname || document.forms[0] || null;
	}

	function controls(name) {
		var frm = form();
		var item = frm && (frm.elements[name] || frm[name]);
		if (!item) {
			item = document.getElementById(name);
		}
		if (!item) {
			return [];
		}
		if (item.length && !item.tagName) {
			return Array.prototype.slice.call(item);
		}
		return [item];
	}

	function field(name) {
		return controls(name)[0] || null;
	}

	function value(name) {
		var item = field(name);
		return item ? item.value : "";
	}

	function setValue(name, data) {
		controls(name).forEach(function (item) {
			item.value = data == null ? "" : String(data);
		});
	}

	function selectText(name) {
		var item = field(name);
		return item && item.options && item.selectedIndex >= 0 ? item.options[item.selectedIndex].text : "";
	}

	function selectValue(name) {
		var item = field(name);
		return item && item.options && item.selectedIndex >= 0 ? item.options[item.selectedIndex].value : "";
	}

	function checked(name) {
		var item = field(name);
		return !!(item && item.checked);
	}

	function setChecked(name, isChecked) {
		var item = field(name);
		if (item) {
			item.checked = !!isChecked;
		}
	}

	function focusSelect(name) {
		var item = field(name);
		if (item) {
			if (typeof item.select === "function") {
				item.select();
			} else if (typeof item.focus === "function") {
				item.focus();
			}
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

	function activityNodes() {
		return childElements(xmlRoot("ActivityData")).filter(function (node) {
			return node.nodeName === "Activity";
		});
	}

	function templateNodes(activity) {
		return childElements(activity).filter(function (node) {
			return node.nodeName === "Template";
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

	function selectedStatus() {
		var radios = controls("radStatus");
		for (var i = 0; i < radios.length; i += 1) {
			if (radios[i].checked) {
				return radios[i].value;
			}
		}
		return radios[1] ? radios[1].value : "I";
	}

	function setStatus(valueToSelect) {
		controls("radStatus").forEach(function (radio) {
			radio.checked = trim(radio.value) === trim(valueToSelect);
		});
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
		xhr.send(serializeXml(islandName));
		return xhr.responseText || "";
	}

	function encodeParam(data) {
		return encodeURIComponent(data == null ? "" : String(data));
	}

	function reloadForActivity(activityNo) {
		var frm = form();
		var passData = "EDT::" + value("hAppCode") + ":" + value("hAppName") + ":" + value("hProcessCode") + ":" + value("hProcessName") + ":" + activityNo;
		frm.action = "ActivityMappingPopUp.asp?PassData=" + encodeParam(passData);
		frm.submit();
	}

	function doneClose() {
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
		window.close();
	}

	function deleteActivity() {
		var response;
		var parts;
		activityNodes().forEach(function (activity) {
			templateNodes(activity).forEach(function (template) {
				var checkbox = field("ChkTempZ" + getAttr(template, "No"));
				if (checkbox && checkbox.checked) {
					setAttr(template, "Del", "Y");
				}
			});
		});
		response = postXml("AppActDel.asp", "ActivityData");
		if (trim(response) === "") {
			return false;
		}
		parts = response.split(":");
		if (parts.length === 2 && trim(parts[0]) === "ActNo") {
			if (trim(parts[1]) !== "0") {
				reloadForActivity(trim(parts[1]));
			} else {
				doneClose();
			}
		} else {
			alert(response);
		}
		return false;
	}

	function editTemplate(tempNo) {
		activityNodes().forEach(function (activity) {
			templateNodes(activity).forEach(function (template) {
				if (String(tempNo) !== getAttr(template, "No")) {
					return;
				}
				setValue("txtTempName", getAttr(template, "Name"));
				setValue("txtTempDesc", getAttr(template, "Description"));
				setValue("txtPath", getAttr(template, "ProgramPath"));
				setStatus(getAttr(template, "Status"));
				Array.prototype.forEach.call(field("selFileType").options, function (option, index) {
					if (trim(option.value).toLowerCase() === trim(getAttr(template, "FileType")).toLowerCase()) {
						field("selFileType").selectedIndex = index;
					}
				});
				setChecked("chkEmail", trim(getAttr(template, "EMAIL")) === "Y");
				setChecked("chkSMS", trim(getAttr(template, "SMS")) === "Y");
			});
		});
		setValue("hTempNo", tempNo);
		return false;
	}

	function clearTable() {
		var table = document.getElementById("tblTemp");
		while (table && table.rows.length > 1) {
			table.deleteRow(1);
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
		var table = document.getElementById("tblTemp");
		var serial = 0;
		var templateCount = 0;
		var row;
		var cell;
		var checkbox;
		var link;
		if (!table) {
			return false;
		}
		clearTable();
		activityNodes().forEach(function (activity) {
			templateNodes(activity).forEach(function (template) {
				var tempNo = getAttr(template, "No");
				serial += 1;
				templateCount += 1;
				row = table.insertRow(table.rows.length);
				appendCell(row, serial, "ExcelSerial", "center");
				cell = appendCell(row, "", "ExcelDisplayCell", "center");
				checkbox = document.createElement("input");
				checkbox.type = "checkbox";
				checkbox.name = "ChkTempZ" + tempNo;
				cell.appendChild(checkbox);
				cell = appendCell(row, "", "ExcelDisplayCell");
				link = document.createElement("a");
				link.href = "#";
				link.className = "ExcelDisplayLink";
				link.title = "Click to Edit Template";
				link.textContent = getAttr(template, "Name");
				link.onclick = function () {
					return editTemplate(tempNo);
				};
				cell.appendChild(link);
				appendCell(row, getAttr(template, "ProgramPath"), "ExcelDisplayCell");
			});
		});
		setValue("hTempCnt", templateCount);
		return true;
	}

	function init() {
		upgradeXml();
		clearTable();
		displayTable();
	}

	function validateTemplateForm(root) {
		if (trim(value("hTempCnt")) === "0") {
			if (trim(value("txtDesc")) === "") {
				alert("Enter Activity Description");
				focusSelect("txtDesc");
				return false;
			}
			if (trim(value("txtTempName")) === "") {
				alert("Enter Template Name");
				focusSelect("txtTempName");
				return false;
			}
			if (field("selFileType").selectedIndex <= 0) {
				alert("Select File Type");
				field("selFileType").focus();
				return false;
			}
			if (trim(value("txtPath")) === "") {
				alert("Enter Program Path");
				focusSelect("txtPath");
				return false;
			}
			return true;
		}
		if (trim(value("txtTempName")) !== "" && trim(value("txtPath")) === "") {
			alert("You have entered Template name so please enter the path for the template");
			focusSelect("txtPath");
			return false;
		}
		if (trim(value("txtPath")) !== "" && trim(value("txtTempName")) === "") {
			alert("You have entered Path for the Template but the Template name is not entered so please enter");
			focusSelect("txtTempName");
			return false;
		}
		if (trim(value("txtTempName")) === "" && trim(value("txtPath")) === "") {
			if (!confirm("Template Details not available you want to update Activity Details?")) {
				return false;
			}
			activityNodes(root).forEach(function (activity) {
				setAttr(activity, "STATUS", selectedStatus());
			});
		}
		return true;
	}

	function populateTemplateAttributes(template) {
		setAttr(template, "Name", value("txtTempName"));
		setAttr(template, "Description", value("txtTempDesc"));
		setAttr(template, "ProgramPath", value("txtPath"));
		setAttr(template, "Status", selectedStatus());
		setAttr(template, "FileType", selectValue("selFileType"));
		setAttr(template, "EMAIL", checked("chkEmail") ? "Y" : "N");
		setAttr(template, "SMS", checked("chkSMS") ? "Y" : "N");
	}

	function updateXmlBeforeSubmit() {
		var doc = xmlDocument("ActivityData");
		var tempNo = value("hTempNo");
		if (!doc || !xmlRoot("ActivityData")) {
			return false;
		}
		if (!validateTemplateForm(xmlRoot("ActivityData"))) {
			return false;
		}
		if (tempNo !== "0") {
			activityNodes().forEach(function (activity) {
				setAttr(activity, "ACTIVITYNAME", value("txtDesc"));
				templateNodes(activity).forEach(function (template) {
					if (tempNo === getAttr(template, "No")) {
						populateTemplateAttributes(template);
					}
				});
			});
			return true;
		}
		activityNodes().forEach(function (activity) {
			var template;
			setAttr(activity, "ACTIVITYNAME", value("txtDesc"));
			template = doc.createElement("Template");
			setAttr(template, "No", "");
			populateTemplateAttributes(template);
			activity.appendChild(template);
		});
		return true;
	}

	function checkSubmit() {
		var response;
		var parts;
		if (!updateXmlBeforeSubmit()) {
			return false;
		}
		response = postXml("AppActivityCreationAndAmendInsert.asp", "ActivityData");
		if (trim(response) === "") {
			return false;
		}
		parts = response.split(":");
		if (parts.length === 2 && trim(parts[0]) === "ActNo") {
			reloadForActivity(trim(parts[1]));
		} else {
			alert(response);
		}
		return false;
	}

	function install() {
		window.DelActivity = deleteActivity;
		window.EditTemplate = editTemplate;
		window.Init = init;
		window.DisplayTable = displayTable;
		window.ClearTable = clearTable;
		window.CheckSubmit = checkSubmit;
		window.DoneClose = doneClose;
	}

	window.ITMSActivityMappingPopupCompat = {
		install: install,
		doneClose: doneClose
	};
}(window, document));
