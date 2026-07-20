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
		return item ? item.value || "" : "";
	}

	function setValue(name, value) {
		var item = field(name);
		if (item) {
			item.value = value == null ? "" : String(value);
		}
	}

	function boolFieldValue(name, fallback) {
		var item = field(name);
		if (!item) {
			return fallback || "N";
		}
		return item.checked ? item.value : fallback || "N";
	}

	function xmlIsland(name) {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
		return window[name] || document[name] || document.getElementById(name) || null;
	}

	function xmlRoot(name) {
		var island = xmlIsland(name);
		return island && island.documentElement || island && island.XMLDocument && island.XMLDocument.documentElement || null;
	}

	function getAttr(node, name) {
		return node && node.getAttribute ? node.getAttribute(name) || "" : "";
	}

	function childElements(node, name) {
		var list = [];
		var wanted = name && String(name).toLowerCase();
		for (var i = 0; node && i < node.childNodes.length; i += 1) {
			if (node.childNodes[i].nodeType === 1 && (!wanted || String(node.childNodes[i].nodeName).toLowerCase() === wanted)) {
				list.push(node.childNodes[i]);
			}
		}
		return list;
	}

	function buildQuery(params) {
		var pairs = [];
		Object.keys(params).forEach(function (name) {
			if (params[name] !== undefined && params[name] !== null) {
				pairs.push(encodeURIComponent(name) + "=" + encodeURIComponent(params[name]));
			}
		});
		return pairs.join("&");
	}

	function selectedText(select) {
		return select && select.selectedIndex >= 0 ? select.options[select.selectedIndex].text : "";
	}

	function itemTypeValue() {
		var selected = trim(valueOf("selIType"));
		if (selected && selected.toUpperCase() !== "SELECT") {
			return selected;
		}
		return window.IType || "";
	}

	function itemTypeName() {
		return selectedText(field("selIType")) || window.ITypeName || "";
	}

	function openDialog(url, args, features, callback) {
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog(url, args, features, callback);
			return;
		}
		var popup = window.open(url, "_blank", "width=650,height=500,resizable=no,scrollbars=yes");
		var timer = window.setInterval(function () {
			if (popup && popup.closed) {
				window.clearInterval(timer);
				callback(popup.returnValue);
			}
		}, 500);
	}

	function splitClassificationReturn(returnData) {
		var parts = trim(returnData).split("*****");
		var codes = [];
		var categories = [];
		var names = [];
		var codeParts;
		var nameParts;
		if (!parts[0] || parts[0] === "-1") {
			return null;
		}
		codeParts = parts[0].split("|");
		for (var i = 0; i < codeParts.length; i += 1) {
			var codeSegment = codeParts[i].split(":");
			if (codeSegment.length > 1) {
				codes.push(codeSegment[codeSegment.length - 1]);
				categories.push(codeSegment[1]);
			} else if (codeSegment[0]) {
				codes.push(codeSegment[0].replace(/^CAT/i, ""));
				categories.push(codeSegment[0].replace(/^CAT/i, ""));
			}
		}
		nameParts = String(parts[1] || "").split("|||");
		for (var j = 0; j < nameParts.length; j += 1) {
			var nameSegment = nameParts[j].split(":");
			if (trim(nameSegment[nameSegment.length - 1])) {
				names.push(nameSegment[nameSegment.length - 1]);
			}
		}
		return {
			codes: codes.join(","),
			categories: categories.join(","),
			names: names.join(",")
		};
	}

	function popClass() {
		var url = "/include/ClassificationSelectPop.asp?" + buildQuery({
			sIType: itemTypeValue(),
			sOrgID: valueOf("hOrgId"),
			sITypename: itemTypeName()
		});
		openDialog(url, "Classification", "dialogHeight:500px;dialogWidth:650px;center:Yes;help:No;resizable:No;status:No", function (returnData) {
			var selected = splitClassificationReturn(returnData);
			var display = document.getElementById("spanClassification") || window.spanClassification;
			if (!selected) {
				return;
			}
			if (display) {
				display.textContent = selected.names;
			}
			setValue("hClassCode", selected.codes);
			Search();
		});
	}

	function selectedEligible() {
		return [
			boolFieldValue("ChkPur"),
			boolFieldValue("ChkSales"),
			boolFieldValue("ChkInv"),
			boolFieldValue("ChkManu")
		].join(",");
	}

	function searchQuery(pageNo) {
		var eligible = selectedEligible();
		var params = {
			PageSize: 15,
			hSelectMode: "M",
			hClassCodes: valueOf("hClassCode"),
			Eligible: eligible.replace(/,/g, ":")
		};
		var searchText = trim(valueOf("txtSearch"));
		var checks = field("chkSearch");
		var itemType = itemTypeValue();
		var cap = "";
		if (checks && checks.length && checks[0].checked) {
			params.SICode = searchText;
		} else if (checks && checks.length && checks[1].checked) {
			params.SIName = searchText;
		}
		if (itemType) {
			params.IType = itemType;
		}
		if (field("radCap") && field("radCap").length) {
			cap = field("radCap")[0].checked ? field("radCap")[0].value : field("radCap")[1].value;
			params.Cap = cap;
		}
		if (pageNo) {
			params.PageNo = pageNo;
		}
		setValue("hEligibleFor", eligible);
		return buildQuery(params);
	}

	function loadXmlIntoIsland(name, text) {
		var island = xmlIsland(name);
		if (island && island.loadXML) {
			island.loadXML(text || "<Root/>");
		}
		return xmlRoot(name);
	}

	function Search(pageNo) {
		var xhr = new XMLHttpRequest();
		xhr.open("GET", "../../Common/XMLGetItemSelectRel.asp?" + searchQuery(pageNo), false);
		xhr.send(null);
		if (xhr.responseText && trim(xhr.responseText)) {
			loadXmlIntoIsland("TempItem", xhr.responseText);
			renderItems();
		} else {
			alert(xhr.responseText || "No response received.");
		}
	}

	function cell(row, text, className, align, html) {
		var td = row.insertCell(-1);
		td.className = className || "ExcelDisplayCell";
		td.align = align || "Left";
		if (html) {
			td.innerHTML = text == null ? "" : String(text);
		} else {
			td.textContent = text == null ? "" : String(text);
		}
		return td;
	}

	function checkboxValue(item, unitId) {
		var name = getAttr(item, "ItemName").replace(/ /g, "-").replace(/"/g, "~~").replace(/'/g, "~");
		return [
			getAttr(item, "ItemCode"),
			getAttr(item, "ClassCode"),
			unitId,
			getAttr(item, "ComItemCode"),
			name,
			"'Y'",
			"0",
			getAttr(item, "AttributeList")
		].join(":");
	}

	function renderItems() {
		var root = xmlRoot("TempItem");
		var items = childElements(root, "Item");
		var table = document.getElementById("tblItem");
		var unitId = valueOf("hOrgId");
		var currPage = root ? getAttr(root, "CurrPage") || "1" : "1";
		var totalPage = root ? getAttr(root, "TotPage") || "1" : "1";
		if (!table) {
			return;
		}
		ClearTable();
		if (!items.length) {
			var emptyRow = table.insertRow(-1);
			var emptyCell = cell(emptyRow, "Data Not Found", "ExcelDisplayCell", "Center");
			emptyCell.colSpan = 8;
			setValue("hPageSelection", "1");
			updatePages(1, 1);
			return;
		}
		for (var i = 0; i < items.length; i += 1) {
			var item = items[i];
			var serialNo = getAttr(item, "SNO") || String(i + 1);
			var itemName = getAttr(item, "ItemName");
			var row = table.insertRow(-1);
			cell(row, serialNo, "ExcelSerial", "Center");
			cell(row, '<input type="checkbox" name="Chkbox' + serialNo + '" value="' + checkboxValue(item, unitId).replace(/"/g, "&quot;") + '">', "ExcelDisplayCell", "center", true);
			cell(row, getAttr(item, "ComItemCode"), "ExcelDisplayCell", "Left");
			if (/ItemListEntryForEdit/i.test(window.location.pathname)) {
				cell(row, '<a href="#" onclick="EditItem(' + serialNo + '); return false;" class="ExcelDisplayLink">' + itemName + '</a>', "ExcelDisplayCell", "Left", true);
			} else {
				cell(row, itemName, "ExcelDisplayCell", "Left");
			}
			cell(row, getAttr(item, "ClassName"), "ExcelDisplayCell", "Left");
			cell(row, getAttr(item, "Stock"), "ExcelDisplayCell", "Right");
			cell(row, getAttr(item, "ItemValue"), "ExcelDisplayCell", "Right");
			cell(row, getAttr(item, "FSNValue"), "ExcelDisplayCell", "Left");
			setValue("hCtr", serialNo);
		}
		updatePages(Number(currPage), Number(totalPage));
		setValue("hPageSelection", currPage);
	}

	function updatePages(currPage, totalPage) {
		var select = field("selPage") || document.getElementById("selPage");
		if (!select || !totalPage) {
			return;
		}
		select.options.length = 0;
		for (var i = 1; i <= totalPage; i += 1) {
			var option = document.createElement("option");
			option.value = i;
			option.text = i === currPage ? "Page " + currPage + " of " + totalPage : "Page " + i;
			option.selected = i === currPage;
			select.add(option);
		}
	}

	function ClearTable() {
		var table = document.getElementById("tblItem");
		if (!table) {
			return;
		}
		while (table.rows.length > 2) {
			table.deleteRow(2);
		}
	}

	function selectedItems() {
		var count = Number(valueOf("hCtr")) || 0;
		var items = [];
		for (var i = 1; i <= count; i += 1) {
			var box = field("Chkbox" + i);
			if (box && box.checked && trim(box.value)) {
				var parts = box.value.split(":");
				items.push({
					itemCode: trim(parts[0]),
					classCode: trim(parts[1]),
					unit: trim(parts[2]),
					category: trim(parts[3]),
					name: trim(parts[4]).replace(/~~/g, "'").replace(/``/g, '"').replace(/~/g, "'"),
					active: trim(parts[5]),
					hold: trim(parts[6]),
					itemTypeId: trim(parts[7])
				});
			}
		}
		return items;
	}

	function saveItemsXml(items, includeItemType) {
		var doc = document.implementation.createDocument("", "Root", null);
		for (var i = 0; i < items.length; i += 1) {
			var node = doc.createElement("Item");
			node.setAttribute("ICode", items[i].itemCode);
			node.setAttribute("CCode", items[i].classCode);
			node.setAttribute("Unit", items[i].unit);
			if (includeItemType && items[i].itemTypeId) {
				node.setAttribute("ItemTypeID", items[i].itemTypeId);
			}
			doc.documentElement.appendChild(node);
		}
		var xhr = new XMLHttpRequest();
		xhr.open("POST", "XMLSave.asp?SessionFlag=True&Value=Inv_ItemDetails_&Folder=Transaction", false);
		xhr.send(new XMLSerializer().serializeToString(doc));
	}

	function firstItem(items) {
		return items.length ? items[0] : {};
	}

	function submitTo(action) {
		form().action = action;
		form().submit();
	}

	function setSingleItemFields(item) {
		setValue("hItemName", item.name || "");
		setValue("hItemCode", item.itemCode || "");
		if (field("selClass")) {
			field("selClass").value = item.classCode || "";
		}
		setValue("hSelectedValue", trim(item.itemCode || "") + "|");
	}

	function EditItem(rowIndex) {
		var box = field("Chkbox" + rowIndex);
		if (!box || !trim(box.value)) {
			return;
		}
		var parts = box.value.split(":");
		setValue("hItemTypeName", "");
		setValue("hItemCode", parts[0]);
		submitTo("itmEditEntry.asp");
	}

	function GotoAction() {
		var choice = trim(valueOf("Choice"));
		var items = selectedItems();
		var item = firstItem(items);
		var itemCodes = items.map(function (entry) { return entry.itemCode; }).join(",");
		if (choice === "SEL") {
			alert("Select Any One Option from the Listbox");
			return;
		}
		if (choice === "ADD") {
			submitTo("ITMCREATIONDEFINITIONENTRY.ASP");
			return;
		}
		if (!choice) {
			return;
		}
		if ((choice === "VEW" || choice === "PAD" || choice === "STR") && items.length !== 1) {
			alert("Select any One Item");
			return;
		}
		if (!items.length && choice !== "ADD") {
			alert("Select a Item");
			return;
		}
		if (choice === "EDT") {
			setValue("hItemTypeName", "");
			setValue("hItemCode", item.itemCode);
			submitTo(trim(valueOf("hEditAction")) === "D" ? "itmEditEntry.asp" : "ItemEditSimple.asp");
			return;
		}
		if (choice === "VEW") {
			window.open("../reports/ItemDetailsDetailsEntry.asp?sTemp=" + [valueOf("hOrgId"), item.classCode, item.itemCode, valueOf("hItemTypeCode")].join(":"), "ItemDetails", "height=540,width=795,toolbar=no,titlebar=no,location=no,directories=no,status=no,personalbar=no,menubar=no,scrollbars=Yes,resizable=yes,top=0,left=0");
			return;
		}
		if (choice === "DEL") {
			setValue("hClassCode", item.classCode);
			setValue("hItemTypeName", "");
			setValue("hItemCode", itemCodes);
			if (confirm("Do U Want To Delete this Item Permanently?")) {
				submitTo("ItmDelete.asp");
			}
			return;
		}
		if (choice === "REC") {
			window.open("../reports/ReceiptItemDetailsEntryNew.asp?sTemp=" + [valueOf("hOrgId"), valueOf("hFromDate"), valueOf("hToDate"), valueOf("hItemTypeCode"), "", "", "I", itemCodes, "SELECTED", "", "", "", ""].join(":"), "ReceiptItem", "height=540,width=795,toolbar=no,titlebar=no,location=no,directories=no,status=no,personalbar=no,menubar=no,scrollbars=No,resizable=yes,top=0,left=0");
			return;
		}
		if (choice === "ISS") {
			window.open("../reports/IssueItemDetailsEntryNew.asp?sTemp=" + [valueOf("hOrgId"), valueOf("hFromDate"), valueOf("hToDate"), valueOf("hItemTypeCode"), "", "", "I", itemCodes, "ALL", "", "ALL", "", "", "", "", "", "", ""].join(":"), "IssuedItems", "height=540,width=795, toolbar=no,titlebar=no,location=no,directories=no,status=no,personalbar=no,menubar=no,scrollbars=No,resizable=yes,top=0,left=0");
			return;
		}
		if (choice === "CON") {
			window.open("../reports/MaterialConsumptionDetailsEntry.asp?sTemp=" + [valueOf("hOrgId"), valueOf("hFromDate"), valueOf("hToDate"), item.classCode, itemCodes, valueOf("hItemTypeCode"), "SELECTED", ""].join(":"), "", "height=540,width=795,toolbar=no,titlebar=no,location=no,directories=no,status=no,personalbar=no,menubar=no,scrollbars=No,resizable=yes,top=0,left=0");
			return;
		}
		if (choice === "CAP" || choice === "NCAP") {
			var toStock = choice === "CAP" ? "N" : "S";
			var message = choice === "CAP" ? "Do you wat to continue to change the Non-Capitalise to Capitalise?" : "Do you wat to continue to change the Capitalise to Non-Capitalise?";
			if (confirm(message)) {
				submitTo("ItmUpdateStockNonStock.asp?ItemCode=" + encodeURIComponent(itemCodes) + "&Stock=" + toStock);
			}
			return;
		}
		if (choice === "STM" || choice === "STR" || choice === "PAD" || choice === "ABS" || choice === "AWS" || choice === "MRG") {
			saveItemsXml(items, choice !== "STM" && choice !== "MRG");
		}
		if (choice === "STM") {
			setSingleItemFields(item);
			setValue("hUnitID", valueOf("hOrgId"));
			submitTo("../transaction/stkMgmtSMEntry.asp");
		} else if (choice === "MRG") {
			setSingleItemFields(item);
			setValue("hUnitID", valueOf("hOrgId"));
			submitTo("../transaction/stkMergeEntry.asp");
		} else if (choice === "ABS") {
			setSingleItemFields(item);
			submitTo("../transaction/stkMgmtArrangeBinEntry.asp");
		} else if (choice === "AWS") {
			setSingleItemFields(item);
			submitTo("../transaction/stkMgmtAttributeWiseStock.asp");
		} else if (choice === "PAD") {
			setSingleItemFields(item);
			submitTo("../transaction/stkMgmtPAEntry.asp");
		} else if (choice === "STR") {
			setSingleItemFields(item);
			submitTo("../transaction/stkMgmtSTEntry.asp");
		} else if (choice === "ACT") {
			setSingleItemFields(item);
			setValue("hItemTypeName", "");
			submitTo(item.active === "N" ? "../transaction/ItemActiveDetails.asp" : "../transaction/ItemInactiveDetails.asp");
		} else if (choice === "HOL") {
			setSingleItemFields(item);
			submitTo(item.hold === "0" ? "../transaction/putItemOnHoldEntry.asp" : "../transaction/reverseItemOnHoldEntry.asp");
		} else if (choice === "VSO") {
			setValue("hItemCode", items.map(function (entry) { return entry.itemCode; }).join(","));
			setValue("hClassCode", items.map(function (entry) { return entry.classCode; }).join(","));
			setValue("hUnitID", items.map(function (entry) { return entry.unit; }).join(","));
			submitTo("../transaction/ItemStockOpening.asp");
		}
	}

	function Sort(nFieldNo, sOrderByField, sOrder) {
		setValue("hField" + trim(nFieldNo), trim(sOrderByField) + ":" + trim(sOrder));
		setValue("hFieldSelected", nFieldNo);
		form().submit();
	}

	function Paginate(pageNo) {
		setValue("hPageSelection", pageNo);
		form().submit();
	}

	function Help() {
		window.open("../HelpFiles/List.htm", "", "toolbar=no,titlebar=no,location=no,directories=no,status=no,menubar=No,scrollbars=yes,resizable=no,width=800px,height=500px,left=10,top=10");
	}

	window.popClass = popClass;
	window.Search = Search;
	window.ClearTable = ClearTable;
	window.EditItem = EditItem;
	window.GotoAction = GotoAction;
	window.Sort = Sort;
	window.Paginate = Paginate;
	window.Help = Help;
}(window, document));
