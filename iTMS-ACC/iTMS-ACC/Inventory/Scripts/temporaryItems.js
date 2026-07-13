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
		var key;
		if (!frm || !frm.elements) {
			return null;
		}
		if (frm.elements[name]) {
			return frm.elements[name];
		}
		for (key in frm.elements) {
			if (Object.prototype.hasOwnProperty.call(frm.elements, key) && key.toLowerCase() === name.toLowerCase()) {
				return frm.elements[key];
			}
		}
		return null;
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

	function xmlObject(name) {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
		return window[name] || document[name] || document.getElementById(name) || null;
	}

	function xmlRoot(name) {
		var object = xmlObject(name);
		return object && object.documentElement || object && object.XMLDocument && object.XMLDocument.documentElement || object && object._doc && object._doc.documentElement || null;
	}

	function Search() {
		form().submit();
		return false;
	}

	function selectedTempItem() {
		var count = Number(valueOf("hCtr")) || 0;
		var selected = "";
		var selectedCount = 0;
		var box;
		for (var i = 1; i <= count; i += 1) {
			box = field("Chkbox" + i);
			if (box && box.checked) {
				selected = box.value;
				selectedCount += 1;
			}
		}
		if (selectedCount === 0) {
			alert("Select an Item");
			return "";
		}
		if (selectedCount > 1) {
			alert("Select any One Item");
			return "";
		}
		setValue("hTempItemCode", selected);
		return selected;
	}

	function DeleteItem() {
		var itemCode = trim(selectedTempItem());
		var xhr;
		if (!itemCode) {
			return false;
		}
		xhr = new XMLHttpRequest();
		xhr.open("POST", "TemporaryItemDeletion.asp?hTempItemCode=" + encodeURIComponent(itemCode), false);
		xhr.send();
		if (trim(xhr.responseText) !== "") {
			alert(xhr.responseText);
		} else {
			alert("Temporary Item Deleted Successfully");
			form().submit();
		}
		return false;
	}

	function ChangeStatus(obj) {
		setValue("hFilterBy", obj ? obj.value : "");
		form().action = "TEMPORARYITEMS.asp";
		form().submit();
		return false;
	}

	function readSelectedItem(returnValue) {
		var root = xmlRoot("ItemSelectData");
		var node;
		if (returnValue && returnValue.getAttribute) {
			node = returnValue;
		}
		if (root && root.hasChildNodes()) {
			node = root.firstElementChild || root.firstChild;
		}
		if (!node || !node.getAttribute) {
			return;
		}
		setValue("hItemCode", node.getAttribute("ItemCode") || "");
		setValue("hClassCode", node.getAttribute("ClassCode") || "");
	}

	function SelectItem(callback) {
		var size = (window.GetWindowSizeForPopup ? window.GetWindowSizeForPopup("1") : "ItemSelectRelPartyCommon.asp:500:850").split(":");
		var programName = size[0] || "ItemSelectRelPartyCommon.asp";
		var popupHeight = size[1] || "500";
		var popupWidth = size[2] || "850";
		var url = "../../Common/" + programName +
			"?orgID=" + encodeURIComponent(valueOf("hOrgID") || valueOf("hOrgId")) +
			"&sIType=" + encodeURIComponent(valueOf("selItemType")) +
			"&Stock=Y&hSelectMode=S&Flag=1&hDispButt=&hDispItem=&CallFrom=";
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog(
				url,
				xmlObject("ItemSelectData"),
				"dialogHeight:" + popupHeight + "px;dialogWidth:" + popupWidth + "px;Status:No",
				function (returnValue) {
					readSelectedItem(returnValue);
					if (typeof callback === "function") {
						callback();
					}
				}
			);
		}
		return false;
	}

	function MapItem() {
		if (!trim(selectedTempItem())) {
			return false;
		}
		SelectItem(function () {
			if (trim(valueOf("hItemCode")) !== "") {
				form().action = "TempItemRelateUpdation.asp";
				form().submit();
			}
		});
		return false;
	}

	function CreateNewItem() {
		form().action = "ITMCREATIONDEFINITIONENTRY.ASP";
		form().submit();
		return false;
	}

	function handleNewTempItem(outDataValue) {
		var parts;
		outDataValue = trim(outDataValue);
		if (!outDataValue || outDataValue === "``") {
			return;
		}
		parts = outDataValue.split("``");
		form().action = "tempItemUpdation.asp?ItemCode=" + encodeURIComponent(parts[0] || "");
		form().submit();
	}

	function NewTempItem() {
		var orgId = valueOf("hOrgID") || valueOf("hOrgId");
		var itemType = valueOf("selItemType") || valueOf("hItemType");
		if (window.CreateTempItem) {
			window.CreateTempItem(2, 2, "Temp Item Creation", itemType, orgId, handleNewTempItem);
		}
		return false;
	}

	window.Search = Search;
	window.DeleteItem = DeleteItem;
	window.ChangeStatus = ChangeStatus;
	window.SelectData = selectedTempItem;
	window.SelectItem = SelectItem;
	window.MapItem = MapItem;
	window.CreateNewItem = CreateNewItem;
	window.NewTempItem = NewTempItem;
}(window, document));
