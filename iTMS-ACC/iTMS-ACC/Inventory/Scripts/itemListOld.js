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
		return item ? item.value : "";
	}

	function setValue(name, value) {
		var item = field(name);
		if (item) {
			item.value = value == null ? "" : String(value);
		}
	}

	function selectedText(name) {
		var select = field(name);
		return select && select.selectedIndex >= 0 && select.options[select.selectedIndex] ? select.options[select.selectedIndex].text : "";
	}

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function xmlObject(name) {
		ensureCompat();
		return window[name] || document[name] || document.getElementById(name) || null;
	}

	function xmlRoot(name) {
		var object = xmlObject(name);
		return object && object.documentElement || object && object.XMLDocument && object.XMLDocument.documentElement || object && object._doc && object._doc.documentElement || null;
	}

	function xmlDocument(name) {
		var object = xmlObject(name);
		var root = xmlRoot(name);
		return object && object.XMLDocument || object && object._doc || root && root.ownerDocument || null;
	}

	function clearXmlRoot(root) {
		while (root && root.firstChild) {
			root.removeChild(root.firstChild);
		}
	}

	function serializeXml(rootOrDoc) {
		var doc = rootOrDoc && rootOrDoc.nodeType === 9 ? rootOrDoc : rootOrDoc && rootOrDoc.ownerDocument;
		return new XMLSerializer().serializeToString(doc || rootOrDoc);
	}

	function selectedItems() {
		var count = Number(valueOf("hCtr")) || 0;
		var items = [];
		var box;
		var parts;
		for (var i = 1; i <= count; i += 1) {
			box = field("Chkbox" + i) || field("chkbox" + i);
			if (box && box.checked && trim(box.value)) {
				parts = box.value.split(":");
				items.push({
					itemCode: trim(parts[0]),
					classCode: trim(parts[1]),
					unit: trim(parts[2]),
					category: trim(parts[3]),
					name: trim(parts[4]).replace(/~~/g, "'").replace(/``/g, '"'),
					active: trim(parts[5]),
					hold: trim(parts[6]),
					itemTypeId: trim(parts[7])
				});
			}
		}
		return items;
	}

	function Search() {
		var eligible = [];
		["ChkPur", "ChkSales", "ChkInv", "ChkManu"].forEach(function (name) {
			var item = field(name);
			if (item && item.checked) {
				eligible.push(item.value);
			}
		});
		setValue("hEligibleFor", eligible.join(","));
		form().submit();
		return false;
	}

	function submitTo(action) {
		form().action = action;
		form().submit();
	}

	function openReport(url, name, features) {
		window.open(url, name || "", features || "height=540,width=795,toolbar=no,titlebar=no,location=no,directories=no,status=no,personalbar=no,menubar=no,scrollbars=No,resizable=yes,top=0,left=0");
	}

	function saveItemsXml(items, includeItemType) {
		var root = xmlRoot("ItemDetails");
		var doc = xmlDocument("ItemDetails");
		var xhr;
		var node;
		if (!root || !doc) {
			return;
		}
		clearXmlRoot(root);
		items.forEach(function (item) {
			node = doc.createElement("Item");
			node.setAttribute("ICode", item.itemCode);
			node.setAttribute("CCode", item.classCode);
			node.setAttribute("Unit", item.unit);
			if (includeItemType && item.itemTypeId) {
				node.setAttribute("ItemTypeID", item.itemTypeId);
			}
			root.appendChild(node);
		});
		xhr = new XMLHttpRequest();
		xhr.open("POST", "XMLSave.asp?SessionFlag=True&Value=Inv_ItemDetails_&Folder=Transaction", false);
		xhr.send(serializeXml(root));
	}

	function setSingleItemFields(item) {
		setValue("hItemName", item.name || "");
		setValue("hItemCode", item.itemCode || "");
		if (field("selClass")) {
			field("selClass").value = item.classCode || "";
		}
		setValue("hSelectedValue", trim(item.itemCode || "") + "|");
	}

	function selectedCodes(items) {
		return items.map(function (item) {
			return item.itemCode;
		}).join(",");
	}

	function GotoAction() {
		var action = trim(valueOf("Choice"));
		var items = selectedItems();
		var item = items[0] || {};
		var itemCodes = selectedCodes(items);
		var tempValues;
		if (action === "SEL") {
			alert("Select Any One Option from the Listbox");
			return false;
		}
		if (action === "ADD") {
			submitTo("ITMCREATIONDEFINITIONENTRY.ASP");
			return false;
		}
		if (!action) {
			return false;
		}
		if (!items.length && action !== "CRW") {
			alert("Select an Item");
			return false;
		}
		if (items.length > 1 && action === "STR") {
			alert("Select any One Item");
			return false;
		}
		if (action === "EDT") {
			if (items.length > 1) {
				alert("Select a Item");
				return false;
			}
			setValue("hItemTypeName", selectedText("selItemType"));
			setValue("hItemCode", item.itemCode);
			submitTo("itmEditEntry.asp");
			return false;
		}
		if (action === "VEW") {
			if (items.length > 1) {
				alert("Select a Item");
				return false;
			}
			tempValues = [valueOf("hOrgId"), item.classCode, item.itemCode, valueOf("hItemTypeCode")].join(":");
			if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
				window.ITMSModernCompat.openModalDialog("../reports/ItemDetailsDetailsEntry.asp?sTemp=" + encodeURIComponent(tempValues), xmlObject("ItemDetails"), "dialogHeight:600px;dialogWidth:800px;center:Yes;help:No;resizable:No;status:No");
			} else {
				openReport("../reports/ItemDetailsDetailsEntry.asp?sTemp=" + encodeURIComponent(tempValues), "ItemDetails", "height=600,width=800,resizable=no,status=no,scrollbars=yes");
			}
			return false;
		}
		if (action === "DEL") {
			setValue("hClassCode", item.classCode);
			setValue("hItemTypeName", selectedText("selItemType"));
			setValue("hItemCode", itemCodes);
			if (confirm("Do U Want To Delete this Item Permanently?")) {
				submitTo("ItmDelete.asp");
			}
			return false;
		}
		if (action === "REC") {
			tempValues = [valueOf("hOrgId"), valueOf("hFromDate"), valueOf("hToDate"), valueOf("hItemTypeCode"), selectedText("selItemType"), item.classCode, "I", itemCodes, item.category, item.category + ":SELECTED"].join(":");
			openReport("../reports/ReceiptItemDetailsEntry.asp?sTemp=" + encodeURIComponent(tempValues), "ReceiptItem");
			return false;
		}
		if (action === "ISS") {
			tempValues = [valueOf("hOrgId"), valueOf("hFromDate"), valueOf("hToDate"), valueOf("hItemTypeCode"), selectedText("selItemType"), item.classCode, "I", itemCodes, item.category, "0", "SELECTED", "0"].join(":");
			openReport("../reports/IssueItemDetailsEntry.asp?sTemp=" + encodeURIComponent(tempValues), "IssuedItems");
			return false;
		}
		if (action === "CON") {
			tempValues = [valueOf("hOrgId"), valueOf("hFromDate"), valueOf("hToDate"), item.classCode, itemCodes, valueOf("hItemTypeCode"), "SELECTED", selectedText("selItemType")].join(":");
			openReport("../reports/MaterialConsumptionDetailsEntry.asp?sTemp=" + encodeURIComponent(tempValues), "");
			return false;
		}
		if (action === "STM" || action === "PAD") {
			if (items.length > 1) {
				saveItemsXml(items, false);
			}
		} else if (action === "STR" || action === "ABS" || action === "AWS") {
			saveItemsXml(items, true);
		} else if (action === "MRG") {
			if (!items.length) {
				alert("Select a Item");
				return false;
			}
			saveItemsXml(items, false);
		} else if (items.length > 1) {
			alert("Select a Item");
			return false;
		}
		if (action === "STM") {
			setSingleItemFields(item);
			setValue("hUnitID", valueOf("hOrgId"));
			submitTo("../transaction/stkMgmtSMEntry.asp");
		} else if (action === "MRG") {
			setSingleItemFields(item);
			setValue("hUnitID", valueOf("hOrgId"));
			submitTo("../transaction/stkMergeEntry.asp");
		} else if (action === "ABS") {
			setSingleItemFields(item);
			submitTo("../transaction/stkMgmtArrangeBinEntry.asp");
		} else if (action === "AWS") {
			setSingleItemFields(item);
			submitTo("../transaction/stkMgmtAttributeWiseStock.asp");
		} else if (action === "PAD") {
			setSingleItemFields(item);
			submitTo("../transaction/stkMgmtPAEntry.asp");
		} else if (action === "STR") {
			setSingleItemFields(item);
			submitTo("../transaction/stkMgmtSTEntry.asp");
		} else if (action === "ACT") {
			setSingleItemFields(item);
			setValue("hItemTypeName", selectedText("selItemType"));
			submitTo(item.active === "N" ? "../transaction/ItemActiveDetails.asp" : "../transaction/ItemInactiveDetails.asp");
		} else if (action === "HOL") {
			setSingleItemFields(item);
			submitTo(item.hold === "0" ? "../transaction/putItemOnHoldEntry.asp" : "../transaction/reverseItemOnHoldEntry.asp");
		}
		return false;
	}

	function Sort(sortBy) {
		setValue("hSortBy", sortBy);
		form().submit();
		return false;
	}

	function Paginate(pageNo) {
		setValue("hPageSelection", pageNo);
		form().submit();
		return false;
	}

	window.Search = Search;
	window.GotoAction = GotoAction;
	window.Sort = Sort;
	window.Paginate = Paginate;
}(window, document));
