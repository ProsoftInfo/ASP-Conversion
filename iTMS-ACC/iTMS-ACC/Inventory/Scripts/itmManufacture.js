(function (window, document) {
	"use strict";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function isAmendPage() {
		return /ItmManufactureAmd\.asp/i.test(window.location.pathname);
	}

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function field(name) {
		var frm = form();
		var item;
		var wanted;
		var i;
		if (!frm || !frm.elements) {
			return null;
		}
		item = frm.elements[name];
		if (item) {
			return item;
		}
		wanted = String(name).toLowerCase();
		for (i = 0; i < frm.elements.length; i += 1) {
			if (String(frm.elements[i].name || "").toLowerCase() === wanted) {
				return frm.elements[i];
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

	function selectedValue(name) {
		var item = field(name);
		return item ? item.value : "";
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

	function xmlRoot(nameOrObject) {
		var object = typeof nameOrObject === "string" ? xmlObject(nameOrObject) : nameOrObject;
		if (window.ITMSModalReturnCompat && window.ITMSModalReturnCompat.xmlRoot) {
			return window.ITMSModalReturnCompat.xmlRoot(object);
		}
		return object && object.documentElement || object && object.XMLDocument && object.XMLDocument.documentElement || object && object._doc && object._doc.documentElement || object && object.nodeType === 1 && object || null;
	}

	function xmlDocument(name) {
		var object = xmlObject(name);
		var root = xmlRoot(object);
		return object && object.XMLDocument || object && object._doc || root && root.ownerDocument || null;
	}

	function childElements(node, name) {
		var result = [];
		var wanted = name ? String(name).toUpperCase() : "";
		var child;
		for (var i = 0; node && i < node.childNodes.length; i += 1) {
			child = node.childNodes[i];
			if (child.nodeType === 1 && (!wanted || String(child.nodeName || "").toUpperCase() === wanted)) {
				result.push(child);
			}
		}
		return result;
	}

	function attr(node, name) {
		return node && node.getAttribute ? node.getAttribute(name) || "" : "";
	}

	function clearNamedChildren(root, name) {
		childElements(root, name).forEach(function (node) {
			root.removeChild(node);
		});
	}

	function serializeXml(rootOrDoc) {
		var doc = rootOrDoc && rootOrDoc.nodeType === 9 ? rootOrDoc : rootOrDoc && rootOrDoc.ownerDocument;
		return new XMLSerializer().serializeToString(doc || rootOrDoc);
	}

	function postXml(url, root) {
		var xhr = new XMLHttpRequest();
		xhr.open("POST", url, false);
		xhr.setRequestHeader("Content-Type", "text/xml");
		xhr.send(serializeXml(root));
		return xhr.responseText || "";
	}

	function selectorSize() {
		var value = typeof window.GetWindowSizeForPopup === "function" ? window.GetWindowSizeForPopup("2") : "PartySelection.asp:500:500";
		var parts = String(value || "PartySelection.asp:500:500").split(":");
		return {
			program: parts[0] || "PartySelection.asp",
			height: parts[1] || "500",
			width: parts[2] || "500"
		};
	}

	function openDialog(url, args, features, callback) {
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			return window.ITMSModernCompat.openModalDialog(url, args, features, callback);
		}
		window.open(url, "_blank");
		return null;
	}

	function applySelectedCustomer(root) {
		childElements(root).forEach(function (node) {
			setValue("txtCustomer", attr(node, "RetField0"));
			setValue("hPartyCode", attr(node, "RetField1"));
		});
	}

	function continueCustomerSelection(returnedValue) {
		var root = xmlRoot(returnedValue);
		var action = String(attr(root, "Action")).toUpperCase();
		var query;
		var size;
		if (!root || action === "CLOSE") {
			return;
		}
		if (action && action !== "DONE") {
			query = attr(root, "PassQuery");
			if (query) {
				size = selectorSize();
				openDialog("../../Common/" + size.program + "?" + query, xmlObject("TempData"), "dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No", continueCustomerSelection);
			}
			return;
		}
		applySelectedCustomer(root);
	}

	function SelectCustomer() {
		var size = selectorSize();
		var url = "../../Common/" + size.program + "?orgID=" + encodeURIComponent(valueOf("hOrgCode"));
		openDialog(url, xmlObject("TempData"), "dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No", continueCustomerSelection);
		return false;
	}

	function CheckBack() {
		window.location.href = (isAmendPage() ? "ItmInvDetAmd.asp" : "ItmInvDet.asp") +
			"?ItemCode=" + encodeURIComponent(valueOf("hItmCode")) +
			"&ClassCode=" + encodeURIComponent(valueOf("hClassCode"));
		return false;
	}

	function createManufactureNode(root) {
		var doc = xmlDocument("OutData") || root.ownerDocument;
		var node = doc.createElement("Manufacture");
		node.setAttribute("partycode", valueOf("hPartyCode"));
		node.setAttribute("alias", valueOf("txtAlias"));
		node.setAttribute("cdn", valueOf("txtCDN"));
		node.setAttribute("mdn", valueOf("txtMDN"));
		node.setAttribute("grade", selectedValue("selGrade"));
		node.setAttribute("modelprocess", selectedValue("selModelProcess"));
		node.setAttribute("noofcavities", valueOf("txtNoofCavities"));
		node.setAttribute("plateno", selectedValue("selMatchPlateNo"));
		node.setAttribute("itemweight", valueOf("txtItemWeight"));
		node.setAttribute("basevalue", valueOf("txtbasevalue"));
		node.setAttribute("itemrate", valueOf("txtitemrate"));
		node.setAttribute("exportrate", valueOf("txtExportRate"));
		node.setAttribute("currency", selectedValue("selCurrency"));
		node.setAttribute("patternmaterial", selectedValue("selPatternMaterial"));
		node.setAttribute("patternowner", selectedValue("selPatternOwner"));
		node.setAttribute("patternavailability", selectedValue("selPatternAvailability"));
		return node;
	}

	function CheckSubmit() {
		var root = xmlRoot("OutData");
		var response;
		if (!isAmendPage() && !trim(valueOf("hItmCode"))) {
			alert("Item is Not Selected");
			return false;
		}
		if (!root) {
			return false;
		}
		root.setAttribute("ItemCode", valueOf("hItmCode"));
		root.setAttribute("ClassCode", valueOf("hClassCode"));
		root.setAttribute("OrgCode", valueOf("hOrgCode"));
		clearNamedChildren(root, "Manufacture");
		root.appendChild(createManufactureNode(root));
		response = trim(postXml("ItmManuInsert.asp", root));
		if (response) {
			alert(response);
		} else {
			alert("Manufacturing Details Stored Successfully.");
			window.location.href = isAmendPage() ? "ITEMLISTENTRYFOREDIT.ASP?ACTN=L" : "ITEMLISTENTRY.ASP?ACTN=L";
		}
		return false;
	}

	window.SelectCustomer = SelectCustomer;
	window.CheckBack = CheckBack;
	window.CheckSubmit = CheckSubmit;
}(window, document));
