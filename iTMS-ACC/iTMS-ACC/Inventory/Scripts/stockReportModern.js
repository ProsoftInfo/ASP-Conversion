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
		return frm && frm.elements ? frm.elements[name] : null;
	}

	function valueOf(name) {
		var item = field(name);
		return item ? item.value || "" : "";
	}

	function setText(id, value) {
		var element = document.getElementById(id);
		if (element) {
			element.textContent = value;
		}
	}

	function xmlIsland(name) {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
		return window[name] || document[name] || document.getElementById(name) || null;
	}

	function xmlRoot(nameOrObject) {
		var item = typeof nameOrObject === "string" ? xmlIsland(nameOrObject) : nameOrObject;
		if (!item) {
			return null;
		}
		return item.documentElement || item.XMLDocument && item.XMLDocument.documentElement || item._doc && item._doc.documentElement || item.nodeType === 1 && item || null;
	}

	function selectNodes(root, expression) {
		if (!root) {
			return [];
		}
		if (root.selectNodes) {
			return root.selectNodes(expression);
		}
		return [];
	}

	function getAttr(node, name) {
		if (!node || !name) {
			return "";
		}
		if (node.getAttribute && node.getAttribute(name) != null) {
			return node.getAttribute(name) || "";
		}
		if (!node.attributes) {
			return "";
		}
		for (var i = 0; i < node.attributes.length; i += 1) {
			if (String(node.attributes[i].name || "").toLowerCase() === String(name).toLowerCase()) {
				return node.attributes[i].value || "";
			}
		}
		return "";
	}

	function parseDate(value) {
		var text = trim(value);
		var match;
		var parsed;
		if (value instanceof Date && !isNaN(value.getTime())) {
			return new Date(value.getFullYear(), value.getMonth(), value.getDate());
		}
		if (!text) {
			return null;
		}
		match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
		if (match) {
			return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
		}
		match = text.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/);
		if (match) {
			var year = Number(match[3]);
			if (year < 100) {
				year += 2000;
			}
			return new Date(year, Number(match[2]) - 1, Number(match[1]));
		}
		parsed = new Date(text);
		return isNaN(parsed.getTime()) ? null : new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
	}

	function compareDates(left, right) {
		var leftDate = parseDate(left);
		var rightDate = parseDate(right);
		if (!leftDate || !rightDate) {
			return 0;
		}
		return leftDate.getTime() - rightDate.getTime();
	}

	function dateControl(name) {
		var control = field(name) || document.getElementById(name);
		if (control && window.ITMSModernCompat) {
			window.ITMSModernCompat.decorateDateInput(control);
		}
		return control;
	}

	function getControlDate(name) {
		var control = dateControl(name);
		if (!control) {
			return "";
		}
		if (typeof control.GetDate === "function") {
			return control.GetDate();
		}
		return control.value || "";
	}

	function setControlDate(name, value) {
		var control = dateControl(name);
		if (control && typeof control.SetDate === "function") {
			control.SetDate(value);
		} else if (control) {
			control.value = value || "";
		}
	}

	function setMinMaxDate(name, minDate, maxDate) {
		var control = dateControl(name);
		if (!control) {
			return;
		}
		if (minDate && typeof control.SetMinDate === "function") {
			control.SetMinDate(minDate);
		}
		if (maxDate && typeof control.SetMaxDate === "function") {
			control.SetMaxDate(maxDate);
		}
	}

	function selectedItemType() {
		var itemType = valueOf("hItemType");
		var select = field("selIType");
		if (select && select.selectedIndex >= 0 && select.value) {
			itemType = select.value;
		}
		return trim(itemType);
	}

	function selectedItemTypeName(itemType) {
		var select = field("selIType");
		if (select && select.selectedIndex >= 0) {
			return select.options[select.selectedIndex].text || "";
		}
		if (itemType === "FIB") {
			return "Fibre";
		}
		if (itemType === "YRN") {
			return "Yarn";
		}
		if (itemType === "STO") {
			return "Stores";
		}
		return "";
	}

	function openModal(url, args, features, callback) {
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog(url, args, features, callback);
			return;
		}
		var popup = window.open(url, "_blank", "width=750,height=500,resizable=no,scrollbars=yes");
		var timer = window.setInterval(function () {
			if (popup && popup.closed) {
				window.clearInterval(timer);
				callback();
			}
		}, 500);
	}

	function buildQuery(values) {
		var pairs = [];
		Object.keys(values).forEach(function (name) {
			if (values[name] != null) {
				pairs.push(encodeURIComponent(name) + "=" + encodeURIComponent(values[name]));
			}
		});
		return pairs.join("&");
	}

	function selectedAttributes() {
		var selects = (form() || document).querySelectorAll('select[name^="selAttrZ"]');
		var values = [];
		for (var i = 0; i < selects.length; i += 1) {
			if (selects[i].selectedIndex > 0 && selects[i].value) {
				values.push(selects[i].value);
			}
		}
		return values.join("?");
	}

	function CheckSubmit(todaysdate) {
		var frm = form();
		var fromDate = getControlDate("ctlFromDate");
		var toDate = getControlDate("ctlToDate");
		var itemType = selectedItemType();
		var itemCode = trim(valueOf("hItemCode")) || "0";
		var classCode = trim(valueOf("hClassCode"));
		var target = /UpdateConsumption/i.test(window.location.pathname) ? "UpdateConsumption.asp" : "StockReplenishment.asp";
		var values;

		if (compareDates(fromDate, todaysdate) > 0) {
			alert("From Date should be less than or equal to Today's Date");
			return;
		}
		if (compareDates(toDate, todaysdate) > 0) {
			alert("To Date should be less than or equal to Today's Date");
			return;
		}
		if (compareDates(fromDate, toDate) > 0) {
			alert("To Date should be greater than or equal to From Date");
			return;
		}

		values = {
			ItemCode: itemCode,
			ClassCode: classCode,
			AttID: selectedAttributes(),
			ItemType: itemType
		};
		if (/UpdateConsumption/i.test(target)) {
			values.FromDate = fromDate;
			values.ToDate = toDate;
		}
		frm.action = target + "?" + buildQuery(values);
		frm.submit();
	}

	function popClass() {
		var orgId = valueOf("hOrgID");
		var itemType = selectedItemType();
		var itemTypeName = selectedItemTypeName(itemType);
		var url = "../../include/ClassificationSelectPop.asp?" + buildQuery({
			sIType: itemType,
			sOrgID: orgId,
			sITypename: itemTypeName
		});
		openModal(url, "Classification", "dialogWidth:650px;dialogHeight:500px;Help:No", function (result) {
			var returnText = trim(result);
			var selected;
			if (!returnText) {
				return;
			}
			selected = returnText.split("*****")[0];
			if (selected === "-1" || !selected) {
				return;
			}
			fetch("../Master/XMLSelectItemClass.asp?" + buildQuery({ sOrgID: orgId, sText: selected }), {
				cache: "no-cache",
				credentials: "same-origin"
			})
				.then(function (response) {
					return response.text();
				})
				.then(function (text) {
					var outData = xmlIsland("OutData");
					var doc = new DOMParser().parseFromString(text || "<Root/>", "text/xml");
					var root = doc.documentElement;
					var names = [];
					var codes = [];
					if (outData && outData.loadXML) {
						outData.loadXML(text);
					}
					for (var i = 0; root && i < root.childNodes.length; i += 1) {
						if (root.childNodes[i].nodeType === 1) {
							names.push(getAttr(root.childNodes[i], "CLASSNAME"));
							codes.push(getAttr(root.childNodes[i], "CLASSCODE"));
						}
					}
					field("hClassName").value = names.join(",");
					field("hClassCode").value = codes.join(",");
					setText("spanClassification", names.join(","));
				})
				.catch(function () {
					alert("Unable to load selected classification details.");
				});
		});
	}

	function clearTable() {
		var table = document.getElementById("tblLot");
		if (!table) {
			return;
		}
		while (table.rows.length) {
			table.deleteRow(0);
		}
	}

	function AddAttrib(count) {
		var table = document.getElementById("tblLot");
		var row;
		var cell;
		var select;
		clearTable();
		if (!table) {
			return;
		}
		row = table.insertRow(-1);
		cell = row.insertCell(-1);
		cell.width = "190";
		cell.className = "ExcelDisplayCell";
		cell.align = "left";
		cell.colSpan = 9;
		for (var i = 1; i <= Number(count || 0); i += 1) {
			select = document.createElement("select");
			select.size = 1;
			select.name = "selAttrZ" + i;
			select.className = "FormElem";
			cell.appendChild(select);
			cell.appendChild(document.createTextNode(" "));
		}
	}

	function addOption(select, text, value) {
		var option = document.createElement("option");
		option.text = text || "";
		option.value = value || "";
		select.add(option);
	}

	function GetData() {
		var root = xmlRoot("Data");
		var attrNodes = selectNodes(root, "//ATTRIBUTES");
		var attId = trim(valueOf("hAttID"));
		AddAttrib(attrNodes.length || 0);
		for (var i = 0; i < attrNodes.length; i += 1) {
			var attrNode = attrNodes[i];
			var attrId = getAttr(attrNode, "ATTRID");
			var select = field("selAttrZ" + (i + 1));
			var groups = selectNodes(attrNode, "./GROUP");
			if (!groups.length) {
				groups = selectNodes(root, '//ATTRIBUTES/GROUP[@ATTRID="' + attrId + '"]');
			}
			if (!select) {
				continue;
			}
			select.options.length = 0;
			addOption(select, getAttr(attrNode, "ATTRNAME"), attrId);
			for (var j = 0; j < groups.length; j += 1) {
				addOption(select, getAttr(groups[j], "OPTIONNAME"), getAttr(groups[j], "OPTIONVALUE"));
				if (select.options[select.options.length - 1].value === attId) {
					select.selectedIndex = select.options.length - 1;
				}
			}
		}
		if (!attId) {
			field("hClassCode").value = "";
			field("hClassName").value = "";
			field("hItemCode").value = "";
			setText("spanClassification", "All Classification");
			setText("idItemName", "All Items");
		}
	}

	function Init(fromDate, toDate) {
		var isUpdate = /UpdateConsumption/i.test(window.location.pathname);
		setControlDate("ctlFromDate", fromDate);
		setControlDate("ctlToDate", isUpdate ? valueOf("hToDate") || toDate : new Date());
		setMinMaxDate("ctlFromDate", fromDate, "");
		setMinMaxDate("ctlToDate", "", toDate);
		GetData();
	}

	function MinDate() {
		var minDate = valueOf("hFrmDate");
		var maxDate = valueOf("hToDate");
		var fromDate = getControlDate("ctlFromDate");
		var toDate = getControlDate("ctlToDate");
		if (compareDates(fromDate, minDate) < 0 || compareDates(fromDate, maxDate) > 0) {
			alert("Date Should be With in the Range " + minDate + " to " + maxDate);
			setControlDate("ctlFromDate", minDate);
			return;
		}
		if (compareDates(toDate, minDate) < 0 || compareDates(toDate, maxDate) > 0) {
			alert("Date Should be With in the Range " + minDate + " to " + maxDate);
			setControlDate("ctlToDate", maxDate);
		}
	}

	function modalRoot(result) {
		if (result && result.nodeType === 1) {
			return result;
		}
		if (result && result.documentElement) {
			return result.documentElement;
		}
		return xmlRoot("ItemData");
	}

	function processSelectedItems() {
		var root = xmlRoot("ItemData");
		var itemCodes = [];
		var itemNames = [];
		if (!root || !root.hasChildNodes()) {
			return;
		}
		for (var i = 0; i < root.childNodes.length; i += 1) {
			var node = root.childNodes[i];
			var attrs;
			var nameValue;
			if (node.nodeType !== 1 || !node.attributes) {
				continue;
			}
			attrs = node.attributes;
			itemCodes.push(attrs.item(2) ? attrs.item(2).nodeValue : getAttr(node, "ItemCode"));
			nameValue = attrs.item(4) ? attrs.item(4).nodeValue : getAttr(node, "ItemDescription");
			itemNames.push(String(nameValue || "").split("--")[0]);
		}
		if (itemCodes.length) {
			field("hItemCode").value = itemCodes.join("|");
			setText("idItemName", itemNames.join(","));
		} else {
			alert("No Items found");
		}
	}

	function openItemDialog(query) {
		var itemData = xmlIsland("ItemData");
		openModal("../../Common/ItemSelectCommon.asp?" + query, itemData, "dialogHeight:500px;dialogWidth:750px;center:Yes;help:No;resizable:No;status:No", function (result) {
			var root = modalRoot(result);
			var action = trim(getAttr(root, "Action")).toUpperCase();
			var nextQuery = trim(getAttr(root, "PassQuery"));
			if (action && action !== "DONE" && action !== "CLOSE" && nextQuery) {
				openItemDialog(nextQuery);
				return;
			}
			if (action !== "CLOSE") {
				if (field("hItemType")) {
					field("hItemType").value = trim(getAttr(root, "ItemType"));
				}
				processSelectedItems();
			}
		});
	}

	function Search() {
		var query = buildQuery({
			orgID: valueOf("hOrgID"),
			sIType: selectedItemType(),
			Stock: "N",
			hSelectMode: "M",
			Flag: window.nFlag || "",
			hClassCodes: valueOf("hClassCode")
		});
		openItemDialog(query);
	}

	function Paginate(pageNo) {
		var frm = form();
		if (field("hPageSelection")) {
			field("hPageSelection").value = pageNo;
		}
		frm.submit();
	}

	window.CheckSubmit = CheckSubmit;
	window.popClass = popClass;
	window.AddAttrib = AddAttrib;
	window.clearTable = clearTable;
	window.GetData = GetData;
	window.Init = Init;
	window.MinDate = MinDate;
	window.Search = Search;
	window.Paginate = Paginate;
}(window, document));
