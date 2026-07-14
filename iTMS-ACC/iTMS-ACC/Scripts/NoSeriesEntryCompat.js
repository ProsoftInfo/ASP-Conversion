(function (window, document) {
	"use strict";

	var popupReturnValue = "";
	var popupReturned = false;

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.formname || document.forms.formname || document.forms[0] || null;
	}

	function field(name) {
		var frm = form();
		return frm && frm.elements ? frm.elements[name] : document.getElementById(name);
	}

	function selectedOption(select) {
		return select && select.options && select.selectedIndex >= 0 ? select.options[select.selectedIndex] : null;
	}

	function selectedText(select) {
		var option = selectedOption(select);
		return option ? option.text : "";
	}

	function selectedValue(select) {
		var option = selectedOption(select);
		return option ? option.value : select && select.value || "";
	}

	function setFieldValue(name, value) {
		var item = field(name);
		if (item) {
			item.value = value == null ? "" : String(value);
		}
	}

	function setText(id, value) {
		var item = document.getElementById(id) || field(id);
		if (item) {
			item.textContent = value == null ? "" : String(value);
		}
	}

	function focusField(name) {
		var item = field(name);
		if (item && item.focus) {
			item.focus();
		}
	}

	function upgradeXml() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.upgradeXmlIslands) {
			window.ITMSModernCompat.upgradeXmlIslands(document);
		}
	}

	function xmlRoot(name) {
		var object;
		upgradeXml();
		object = window[name] || document[name] || document.getElementById(name);
		return object && object.documentElement || object && object.XMLDocument && object.XMLDocument.documentElement || object && object._doc && object._doc.documentElement || null;
	}

	function childElements(node) {
		return Array.prototype.slice.call(node && node.childNodes || []).filter(function (child) {
			return child.nodeType === 1;
		});
	}

	function attrByIndex(node, index) {
		return node && node.attributes && node.attributes[index] ? node.attributes[index].nodeValue : "";
	}

	function attr(node, name) {
		return node && node.getAttribute ? node.getAttribute(name) || "" : "";
	}

	function clearTable() {
		var table = document.getElementById("tblBook");
		var row;
		if (!table) {
			return;
		}
		while (table.rows.length) {
			table.deleteRow(0);
		}
		row = table.insertRow(0);
		window.InsertCell(row, 1, "", "S.No", "ExcelSerial", "Center", "", 0, 0, 0, 0, "");
		window.InsertCell(row, 1, "", "Period", "ExcelHeaderCell", "left", "", 0, 0, 0, 0, "");
		window.InsertCell(row, 1, "", "StartNo", "ExcelHeaderCell", "left", "", 0, 0, 0, 0, "");
		window.InsertCell(row, 1, "", "Prefix", "ExcelHeaderCell", "left", "", 0, 0, 0, 0, "");
		window.InsertCell(row, 1, "", "Suffix", "ExcelHeaderCell", "left", "", 0, 0, 0, 0, "");
	}

	function periodLabel(periodType, period) {
		if (periodType === "M") {
			return "Month-" + period;
		}
		if (periodType === "Q") {
			return "Quater-" + period;
		}
		if (periodType === "Y") {
			return "Yearly";
		}
		return period;
	}

	function displayTable() {
		var root;
		var seriesNo;
		var rowIndex = 1;
		var select = field("selNoSeries");
		clearTable();
		if (field("selUnit") && field("selUnit").selectedIndex === 0) {
			alert("Select Unit");
			focusField("selUnit");
			return false;
		}
		if (field("selActType") && field("selActType").selectedIndex === 0) {
			alert("Select Activity");
			focusField("selActType");
			return false;
		}
		if (!select || select.selectedIndex === 0) {
			return false;
		}
		root = xmlRoot("SeriesNoData");
		seriesNo = select.value;
		childElements(root).forEach(function (header) {
			if (attrByIndex(header, 0) !== seriesNo) {
				return;
			}
			setFieldValue("hSeriesType", attrByIndex(header, 2));
			setFieldValue("hSeriesLen", attrByIndex(header, 4));
			childElements(header).forEach(function (entry) {
				var row = document.getElementById("tblBook").insertRow(rowIndex);
				var entryNo = attrByIndex(entry, 0);
				var period = periodLabel(attrByIndex(header, 3), attrByIndex(entry, 1));
				window.InsertCell(row, 1, "", rowIndex, "ExcelSerial", "Center", "", 0, 0, 0, 0, "");
				window.InsertCell(row, 1, "", period, "ExcelDisplayCell", "left", "", 0, 0, 0, 0, "");
				window.InsertCell(row, 2, "txtStartNo" + entryNo, attrByIndex(entry, 2), "ExcelInputCell", "", "", 5, 4, 0, 0, "");
				window.InsertCell(row, 2, "txtPrefix" + entryNo, attrByIndex(entry, 3), "ExcelInputCell", "", "", 11, 10, 0, 0, "");
				window.InsertCell(row, 2, "txtSuffix" + entryNo, attrByIndex(entry, 4), "ExcelInputCell", "", "", 11, 10, 0, 0, "");
				rowIndex += 1;
			});
		});
		return false;
	}

	function popSeriesNo() {
		var select = field("selNoSeries");
		var root = xmlRoot("SeriesNoData");
		if (!select) {
			return false;
		}
		select.options.length = 0;
		select.options[select.options.length] = new Option("Select Number Series", "0");
		childElements(root).forEach(function (header) {
			if (attrByIndex(header, 3) === "M") {
				select.options[select.options.length] = new Option(attrByIndex(header, 1), attrByIndex(header, 0));
			}
		});
		return false;
	}

	function clearAll() {
		if (field("selActType")) {
			field("selActType").selectedIndex = 0;
		}
		if (field("selUnit")) {
			field("selUnit").selectedIndex = 0;
		}
		if (field("selNoSeries")) {
			field("selNoSeries").selectedIndex = 0;
		}
		clearTable();
		return false;
	}

	function validateForm() {
		if (field("selUnit") && field("selUnit").selectedIndex === 0) {
			alert("Select Unit");
			focusField("selUnit");
			return false;
		}
		if (field("selActType") && field("selActType").selectedIndex === 0) {
			alert("Select Activity");
			focusField("selActType");
			return false;
		}
		if (field("selNoSeries") && field("selNoSeries").selectedIndex === 0) {
			alert("Select Number Series");
			focusField("selNoSeries");
			return false;
		}
		setFieldValue("hActivityName", selectedText(field("selActType")));
		form().submit();
		return false;
	}

	function optclick(obj) {
		setFieldValue("hOptval", obj && obj.value);
		return false;
	}

	function parseClassificationReturn(returnData) {
		var parts = String(returnData || "").split("*****");
		var classCodes = [];
		var categoryCodes = [];
		var classNames = [];
		if (parts[0] === "-1" || trim(parts[0]) === "") {
			return null;
		}
		parts[0].split("|").forEach(function (value) {
			var split = value.split(":");
			if (split.length > 1) {
				classCodes.push(split[split.length - 1]);
				categoryCodes.push(split[1]);
			} else if (split[0]) {
				categoryCodes.push(split[0].substring(3));
			}
		});
		String(parts[1] || "").split("|||").forEach(function (value) {
			var split = value.split(":");
			if (trim(split[split.length - 1]) !== "") {
				classNames.push(split[split.length - 1]);
			}
		});
		return {
			classCode: classCodes.join(","),
			categoryCode: categoryCodes.join(","),
			className: classNames.join(",")
		};
	}

	function applyClassification(returnData) {
		var result = parseClassificationReturn(returnData);
		var category = field("selCategory");
		if (!result) {
			return;
		}
		setText("txtClass", result.className);
		setFieldValue("hClassCode", result.classCode);
		setFieldValue("hCatCode", result.categoryCode);
		if (category) {
			for (var i = 0; i < category.options.length; i += 1) {
				if (trim(category.options[i].value) === trim(result.categoryCode)) {
					category.selectedIndex = i;
					category.disabled = true;
					break;
				}
			}
		}
	}

	function SelectClassifcation() {
		var orgId = selectedValue(field("selUnit"));
		var url = "/include/ClassificationSelectPop.asp?sIType=1&sOrgID=" + encodeURIComponent(orgId) + "&sITypename=&SelMode=M";
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog(url, "Classification", "dialogHeight:500px;dialogWidth:650px;center:Yes;help:No;resizable:No;status:No", applyClassification);
		} else {
			alert("Popup support is not available.");
		}
		return false;
	}

	function selectedOptions(select) {
		return Array.prototype.slice.call(select && select.options || []).filter(function (option) {
			return option.selected;
		});
	}

	function returnPopupValue(value) {
		popupReturnValue = value || "";
		popupReturned = true;
		if (window.ITMSModalReturnCompat && window.ITMSModalReturnCompat.returnAndClose) {
			window.ITMSModalReturnCompat.returnAndClose(popupReturnValue);
		} else {
			window.returnValue = popupReturnValue;
			window.close();
		}
	}

	function buildSelectionReturn(select, includeShortName) {
		var ids = [];
		var names = [];
		var shortNames = [];
		selectedOptions(select).forEach(function (option) {
			var parts = String(option.value || "").split(":");
			ids.push(parts[0] || "");
			names.push(option.text || "");
			if (includeShortName) {
				shortNames.push(parts[1] || "");
			}
		});
		return includeShortName ? ids.join(":") + "``" + names.join(":") + "``" + shortNames.join(":") : ids.join(":") + "``" + names.join(":");
	}

	function checksubmit() {
		var itemSelect = field("selPartyHead");
		var workGroupSelect = field("SelWorkGroup");
		if (itemSelect) {
			if (itemSelect.selectedIndex < 0) {
				alert("Select Party Type");
				itemSelect.focus();
				return false;
			}
			returnPopupValue(buildSelectionReturn(itemSelect, field("hCallTy") && field("hCallTy").value === "A"));
			return false;
		}
		if (workGroupSelect) {
			if (workGroupSelect.selectedIndex < 0) {
				alert("Select Work Group");
				workGroupSelect.focus();
				return false;
			}
			returnPopupValue(buildSelectionReturn(workGroupSelect, false));
		}
		return false;
	}

	function finalcancel() {
		if (!popupReturned && (field("selPartyHead") || field("SelWorkGroup"))) {
			if (window.ITMSModalReturnCompat && window.ITMSModalReturnCompat.returnValue) {
				window.ITMSModalReturnCompat.returnValue(popupReturnValue);
			} else {
				window.returnValue = popupReturnValue;
			}
		}
		return false;
	}

	function installPopupKeys() {
		document.addEventListener("keydown", function (event) {
			if (event.key === "Escape" || event.keyCode === 27) {
				finalcancel();
				window.close();
			}
		});
		window.addEventListener("beforeunload", finalcancel);
	}

	window.Activity_change = function () {};
	window.Optclick = optclick;
	window.DisplayTable = displayTable;
	window.popSeriesNo = popSeriesNo;
	window.ClearTable = clearTable;
	window.ClearAll = clearAll;
	window.validateForm = validateForm;
	window.SelectClassifcation = SelectClassifcation;
	window.checksubmit = checksubmit;
	window.finalcancel = finalcancel;
	window.window_onunload = finalcancel;
	window.document_onkeypress = function () {};
	installPopupKeys();
}(window, document));
