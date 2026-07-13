(function (window, document) {
	"use strict";

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function xmlIsland(name) {
		ensureCompat();
		return window[name] || document[name] || null;
	}

	function xmlDocument(name) {
		var island = xmlIsland(name);
		return island && (island.XMLDocument || island._doc || island);
	}

	function xmlRoot(name) {
		var island = xmlIsland(name);
		return island && (island.documentElement || island.XMLDocument && island.XMLDocument.documentElement || island._doc && island._doc.documentElement);
	}

	function selectedRadioValue(radios) {
		if (!radios) {
			return "";
		}
		if (radios.length == null) {
			return radios.checked ? radios.value : "";
		}
		for (var i = 0; i < radios.length; i += 1) {
			if (radios[i].checked) {
				return radios[i].value;
			}
		}
		return "";
	}

	function clearChildren(node, tagName) {
		var children = Array.prototype.slice.call(node.childNodes || []);
		children.forEach(function (child) {
			if (!tagName || String(child.nodeName).toLowerCase() === String(tagName).toLowerCase()) {
				node.removeChild(child);
			}
		});
	}

	function addAttribute(node, name, value) {
		node.setAttribute(name, value == null ? "" : value);
	}

	function CheckSubmit() {
		var frm = form();
		var outDoc = xmlDocument("OutData");
		var root = xmlRoot("OutData");
		var inv;
		var request;
		var nextPage = window.ITMS_INV_DETAIL_NEXT || "ItmManufacture.asp";
		if (window.ITMS_INV_DETAIL_REQUIRE_ITEM && String(frm.hItmCode.value || "").replace(/^\s+|\s+$/g, "") === "") {
			alert("Item is Not Selected");
			return false;
		}
		if (!outDoc || !root) {
			alert("Unable to prepare inventory details.");
			return false;
		}

		addAttribute(root, "ItemCode", frm.hItmCode.value);
		addAttribute(root, "ClassCode", frm.hClassCode.value);
		addAttribute(root, "OrgCode", frm.hOrgCode.value);
		clearChildren(root, "Inventory");

		inv = outDoc.createElement("Inventory");
		root.appendChild(inv);
		addAttribute(inv, "ABC", selectedRadioValue(frm.radABC));
		addAttribute(inv, "FSN", "");
		addAttribute(inv, "VED", selectedRadioValue(frm.radVED));
		addAttribute(inv, "ACC", "");
		addAttribute(inv, "Fast", frm.txtFastMovCriteria ? frm.txtFastMovCriteria.value : "");
		addAttribute(inv, "Slow", frm.txtSlowMovCriteria ? frm.txtSlowMovCriteria.value : "");
		addAttribute(inv, "Non", frm.txtNonMovCriteria ? frm.txtNonMovCriteria.value : "");
		if (frm.txtReLvl) {
			addAttribute(inv, "RL", frm.txtReLvl.value);
		}
		if (frm.txtReQty) {
			addAttribute(inv, "RQ", frm.txtReQty.value);
		}
		if (frm.txtEcQty) {
			addAttribute(inv, "EQ", frm.txtEcQty.value);
		}

		request = new XMLHttpRequest();
		request.open("POST", "ItemInvInsert.asp", false);
		request.send(outDoc);
		if (String(request.responseText || "").replace(/^\s+|\s+$/g, "") !== "") {
			alert(request.responseText);
		} else {
			alert("Inventory Details Stored Successfully.");
			window.location.href = nextPage + "?ItemCode=" + encodeURIComponent(frm.hItmCode.value) + "&ClassCode=" + encodeURIComponent(frm.hClassCode.value);
		}
		return true;
	}

	function LoadDraftedDetails() {
		ensureCompat();
	}

	function setSelected(select, value) {
		var text = String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
		if (!select || !select.options) {
			return;
		}
		for (var i = 0; i < select.options.length; i += 1) {
			if (String(select.options[i].value).replace(/^\s+|\s+$/g, "") === text) {
				select.options[i].selected = true;
				return;
			}
		}
	}

	function CheckBack() {
		var frm = form();
		var backPage = window.ITMS_INV_DETAIL_BACK || "ItmDetailedDefn.asp";
		window.location.href = backPage + "?ItemCode=" + encodeURIComponent(frm.hItmCode.value) + "&ClassCode=" + encodeURIComponent(frm.hClassCode.value);
	}

	window.CheckSubmit = CheckSubmit;
	window.LoadDraftedDetails = LoadDraftedDetails;
	window.setSelected = setSelected;
	window.CheckBack = CheckBack;
}(window, document));
