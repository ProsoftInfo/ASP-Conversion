(function (window, document) {
	"use strict";

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function form() {
		return document.forms.formname || document.forms[0] || null;
	}

	function field(name) {
		var frm = form();
		return frm && frm.elements && frm.elements[name] || document.getElementById(name) || null;
	}

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function valueOf(name) {
		var item = field(name);
		return trim(item && "value" in item ? item.value : item && item.textContent);
	}

	function setValue(name, value) {
		var item = field(name);
		if (item && "value" in item) {
			item.value = value == null ? "" : String(value);
		}
	}

	function setText(id, value) {
		var item = document.getElementById(id);
		if (item) {
			item.textContent = value == null ? "" : String(value);
		}
	}

	function selectedValue(nameOrSelect) {
		var select = typeof nameOrSelect === "string" ? field(nameOrSelect) : nameOrSelect;
		if (!select || !select.options || select.selectedIndex < 0) {
			return "";
		}
		return trim(select.options[select.selectedIndex].value);
	}

	function radioValue(name) {
		var items = field(name);
		if (!items) {
			return "";
		}
		if (items.length) {
			for (var i = 0; i < items.length; i += 1) {
				if (items[i].checked) {
					return items[i].value;
				}
			}
			return "";
		}
		return items.checked ? items.value : "";
	}

	function pad2(value) {
		return value < 10 ? "0" + value : String(value);
	}

	function parseDate(value) {
		var text = trim(value);
		var match;
		var year;
		if (!text) {
			return null;
		}
		match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
		if (match) {
			return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
		}
		match = text.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/);
		if (match) {
			year = Number(match[3]);
			if (year < 100) {
				year += 2000;
			}
			return new Date(year, Number(match[2]) - 1, Number(match[1]));
		}
		return null;
	}

	function dateDiffDays(start, end) {
		var startDate = parseDate(start);
		var endDate = parseDate(end);
		if (!startDate || !endDate) {
			return 0;
		}
		return Math.floor((endDate.getTime() - startDate.getTime()) / 86400000);
	}

	function dateValue(name) {
		var item = field(name);
		if (!item) {
			return "";
		}
		if (typeof item.GetDate === "function") {
			return trim(item.GetDate());
		}
		if (typeof item.getDate === "function") {
			return trim(item.getDate());
		}
		return trim(item.value);
	}

	function setDateValue(name, value) {
		var item = field(name);
		if (!item) {
			return;
		}
		if (typeof item.SetDate === "function") {
			item.SetDate(value);
		} else if (typeof item.setDate === "function") {
			item.setDate(value);
		} else {
			item.value = value;
		}
	}

	function setDateMin(name, value) {
		var item = field(name);
		if (item && typeof item.SetMinDate === "function") {
			item.SetMinDate(value);
		}
	}

	function setDateMax(name, value) {
		var item = field(name);
		if (item && typeof item.SetMaxDate === "function") {
			item.SetMaxDate(value);
		}
	}

	function xmlIsland(name) {
		ensureCompat();
		return window[name] || document[name] || document.getElementById(name);
	}

	function xmlRoot(value) {
		var target = typeof value === "string" ? xmlIsland(value) : value;
		return target && (target.documentElement || target.XMLDocument && target.XMLDocument.documentElement || target._doc && target._doc.documentElement) || null;
	}

	function elementChildren(node, name) {
		var result = [];
		var wanted = name && String(name).toLowerCase();
		var children = node && node.childNodes || [];
		for (var i = 0; i < children.length; i += 1) {
			if (children[i].nodeType === 1 && (!wanted || String(children[i].nodeName).toLowerCase() === wanted)) {
				result.push(children[i]);
			}
		}
		return result;
	}

	function getAttr(node, name) {
		return trim(node && node.getAttribute ? node.getAttribute(name) : "");
	}

	function attrAny(node, names) {
		for (var i = 0; i < names.length; i += 1) {
			var value = getAttr(node, names[i]);
			if (value !== "") {
				return value;
			}
		}
		return "";
	}

	function buildQuery(values) {
		var parts = [];
		Object.keys(values).forEach(function (key) {
			if (values[key] != null) {
				parts.push(encodeURIComponent(key) + "=" + encodeURIComponent(values[key]));
			}
		});
		return parts.join("&");
	}

	function syncGet(url) {
		var xhr = new XMLHttpRequest();
		xhr.open("GET", url, false);
		xhr.send(null);
		return xhr;
	}

	function openModal(url, args, features, callback) {
		ensureCompat();
		if (!window.ITMSModernCompat || !window.ITMSModernCompat.openModalDialog) {
			alert("Modern browser compatibility script is still loading. Please try again.");
			return null;
		}
		return window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
	}

	function itemTypeValue() {
		return valueOf("hItemType") || selectedValue("selIType");
	}

	function itemTypeName(itemType) {
		var type = trim(itemType).toUpperCase();
		if (type === "FIB") {
			return "Fibre";
		}
		if (type === "YRN") {
			return "Yarn";
		}
		if (type === "STO") {
			return "Stores";
		}
		return "";
	}

	function selectedCategory() {
		var select = field("selCategory");
		if (!select || select.selectedIndex < 0) {
			return "";
		}
		return selectedValue(select);
	}

	function navigationQuery(extra) {
		var values = {
			ItemCode: valueOf("hItemCode") || "0",
			ClassCode: valueOf("hClassCode"),
			AttID: valueOf("hAttID"),
			ItemType: itemTypeValue(),
			FromDate: dateValue("ctlFromDate") || valueOf("hFrmDate"),
			ToDate: dateValue("ctlToDate") || valueOf("hToDate"),
			Category: valueOf("hCategory"),
			PartyCode: valueOf("hPartyCode"),
			IMType: radioValue("radType") || "I"
		};
		Object.keys(extra || {}).forEach(function (key) {
			values[key] = extra[key];
		});
		return values;
	}

	function submitToList(extra) {
		form().action = "MaterialConsumptions.asp?" + buildQuery(navigationQuery(extra));
		form().submit();
	}

	function parseClassificationReturn(value) {
		var text = trim(value);
		var selected;
		if (!text || text === "-1") {
			return "";
		}
		selected = text.split("*****")[0];
		return selected === "-1" ? "" : selected;
	}

	function loadSelectedClassNames(orgId, selected) {
		var xhr = syncGet("../Master/XMLSelectItemClass.asp?" + buildQuery({ sOrgID: orgId, sText: selected }));
		var outData = xmlIsland("OutData");
		var doc;
		var root;
		var names = [];
		var codes = [];
		if (outData && typeof outData.loadXML === "function") {
			outData.loadXML(xhr.responseText || "<Root/>");
			root = xmlRoot(outData);
		} else {
			doc = new DOMParser().parseFromString(xhr.responseText || "<Root/>", "text/xml");
			root = doc.documentElement;
		}
		elementChildren(root).forEach(function (node) {
			names.push(getAttr(node, "CLASSNAME"));
			codes.push(getAttr(node, "CLASSCODE"));
		});
		setValue("hClassName", names.join(","));
		setValue("hClassCode", codes.join(","));
		setText("spanClassification", names.join(",") || "All Classification");
	}

	window.ListChange = function () {
		submitToList({ IMType: radioValue("radType") || "I" });
		return false;
	};

	window.MatConsumption = function (itemCode, classCode, fromDate, toDate) {
		form().action = "MatConsDetailsEntry.asp?" + buildQuery({ ItemCode: itemCode, ClassCode: classCode, FromDate: fromDate, ToDate: toDate });
		form().submit();
		return false;
	};

	window.MatConsumptionIss = function (issueNo, fromDate, toDate) {
		form().action = "MatConsDetailsEntry.asp?" + buildQuery({ FromDate: fromDate, ToDate: toDate, IssueNo: issueNo });
		form().submit();
		return false;
	};

	window.MatConsumptionIssueWise = function (issueNo) {
		form().action = "MatConsDetailsEntry.asp?" + buildQuery({ IssueNo: issueNo });
		form().submit();
		return false;
	};

	window.Paginate = function (page) {
		if (!page || String(page) === "0") {
			return false;
		}
		setValue("hPageSelection", page);
		form().action = "MaterialConsumptions.asp?" + buildQuery(navigationQuery());
		form().submit();
		return false;
	};

	window.CheckSubmit = function (todaysDate) {
		var fromDate = dateValue("ctlFromDate");
		var toDate = dateValue("ctlToDate");
		if (dateDiffDays(fromDate, todaysDate) < 0) {
			alert("From Date should be less than or equal to Today's Date");
			return false;
		}
		if (dateDiffDays(toDate, todaysDate) < 0) {
			alert("To Date should be less than or equal to Today's Date");
			return false;
		}
		if (dateDiffDays(fromDate, toDate) < 0) {
			alert("To Date should be greater than or equal to From Date");
			return false;
		}
		setValue("hCategory", selectedCategory());
		submitToList();
		return false;
	};

	window.popClass = function () {
		var orgId = valueOf("hOrgID");
		var itemType = itemTypeValue();
		var url = "../../include/ClassificationSelectPop.asp?" + buildQuery({
			sIType: itemType,
			sOrgID: orgId,
			sITypename: itemTypeName(itemType),
			SelMode: "M"
		});
		openModal(url, "Classification", "dialogWidth:650px;dialogHeight:500px;Help:No", function (result) {
			var selected = parseClassificationReturn(result);
			if (selected) {
				loadSelectedClassNames(orgId, selected);
			}
		});
		return false;
	};

	window.AddAttrib = function () {};

	window.clearTable = function () {
		var table = document.getElementById("tblLot");
		if (!table) {
			return false;
		}
		while (table.rows.length) {
			table.deleteRow(0);
		}
		return false;
	};

	window.GetData = function () {
		var category = valueOf("hCategory");
		var select = field("selCategory");
		if (select && category) {
			for (var i = 0; i < select.options.length; i += 1) {
				if (trim(select.options[i].value) === category) {
					select.selectedIndex = i;
					break;
				}
			}
		}
		if (!valueOf("hAttID")) {
			setValue("hClassCode", "");
			setValue("hClassName", "");
			setValue("hItemCode", "");
			setText("spanClassification", "All Classification");
			setText("idItemName", "All Items");
		}
		return false;
	};

	window.Init = function (fromDate, toDate) {
		ensureCompat();
		setDateMin("ctlFromDate", fromDate);
		setDateMax("ctlToDate", toDate);
		if (!dateValue("ctlFromDate")) {
			setDateValue("ctlFromDate", fromDate);
		}
		if (!dateValue("ctlToDate")) {
			setDateValue("ctlToDate", toDate);
		}
		window.GetData();
		return false;
	};

	window.MinDate = function () {
		var min = valueOf("hFrmDate");
		var max = valueOf("hToDate");
		var fromDate = dateValue("ctlFromDate");
		var toDate = dateValue("ctlToDate");
		if (dateDiffDays(min, fromDate) < 0 || dateDiffDays(fromDate, max) < 0) {
			alert("Date Should be With in the Range " + min + " to " + max);
			setDateValue("ctlFromDate", min);
			return false;
		}
		if (dateDiffDays(min, toDate) < 0 || dateDiffDays(toDate, max) < 0) {
			alert("Date Should be With in the Range " + min + " to " + max);
			setDateValue("ctlToDate", max);
			return false;
		}
		return false;
	};

	window.Search = function () {
		var size = typeof window.GetWindowSizeForPopup === "function" ? String(window.GetWindowSizeForPopup("1")).split(":") : ["ItemSelectCommon.asp", "500", "750"];
		var url = "../../Common/" + (size[0] || "ItemSelectCommon.asp") + "?" + buildQuery({
			orgID: valueOf("hOrgID"),
			sIType: itemTypeValue(),
			Stock: "N",
			hSelectMode: "M",
			Flag: "1",
			hClassCodes: valueOf("hClassCode")
		});
		openModal(url, xmlIsland("ItemData"), "dialogHeight:" + (size[1] || "500") + "px;dialogWidth:" + (size[2] || "750") + "px;Status:No", function (result) {
			var root = xmlRoot(result) || xmlRoot("ItemData");
			var codes = [];
			var names = [];
			elementChildren(root).forEach(function (node) {
				var code = attrAny(node, ["ItemCode", "ITEMCODE", "RetField1"]);
				var name = attrAny(node, ["ItemName", "ITEMNAME", "RetField4"]);
				if (code) {
					codes.push(code);
				}
				if (name) {
					names.push(name.split("--")[0]);
				}
			});
			if (!codes.length) {
				alert("No Items found");
				return;
			}
			setValue("hItemCode", codes.join(","));
			setText("idItemName", names.join(","));
		});
		return false;
	};

	window.SelectParty = function () {
		var size = typeof window.GetWindowSizeForPopup === "function" ? String(window.GetWindowSizeForPopup("2")).split(":") : ["PartySelection.asp", "500", "500"];
		var url = "../../Common/" + (size[0] || "PartySelection.asp") + "?" + buildQuery({ orgID: valueOf("hOrgID") });
		openModal(url, xmlIsland("PartyData"), "dialogHeight:" + (size[1] || "500") + "px;dialogWidth:" + (size[2] || "500") + "px;Status:No", function (result) {
			var root = xmlRoot(result) || xmlRoot("PartyData");
			var codes = [];
			var names = [];
			elementChildren(root).forEach(function (node) {
				var code = getAttr(node, "RetField1");
				var name = getAttr(node, "RetField0");
				if (code) {
					codes.push(code);
				}
				if (name) {
					names.push(name.split("--")[0]);
				}
			});
			if (!codes.length) {
				alert("No Items found");
				return;
			}
			setValue("hPartyCode", codes.join(","));
			setText("spanParty", names.join(","));
		});
		return false;
	};
}(window, document));
