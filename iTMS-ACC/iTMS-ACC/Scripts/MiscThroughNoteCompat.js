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
		var lowerName;
		var i;
		if (!frm || !name) {
			return null;
		}
		if (frm.elements[name] || frm[name]) {
			return frm.elements[name] || frm[name];
		}
		lowerName = String(name).toLowerCase();
		for (i = 0; i < frm.elements.length; i += 1) {
			if (String(frm.elements[i].name || "").toLowerCase() === lowerName) {
				return frm.elements[i];
			}
		}
		return null;
	}

	function firstField(item) {
		return item && item.length && !item.tagName ? item[0] : item;
	}

	function valueOf(name, fallback) {
		var item = field(name);
		var i;
		if (!item) {
			return fallback == null ? "" : fallback;
		}
		if (item.length && !item.tagName) {
			for (i = 0; i < item.length; i += 1) {
				if (item[i].checked) {
					return item[i].value;
				}
			}
			return item[0] ? item[0].value : "";
		}
		return item.value == null ? "" : item.value;
	}

	function setValue(name, value) {
		var item = firstField(field(name));
		if (item) {
			item.value = value == null ? "" : String(value);
		}
	}

	function byId(id) {
		return document.getElementById(id);
	}

	function setHtml(id, html) {
		var item = byId(id);
		if (item) {
			item.innerHTML = html == null ? "" : String(html);
		}
	}

	function ready(callback) {
		if (document.readyState === "loading") {
			document.addEventListener("DOMContentLoaded", callback);
		} else {
			callback();
		}
	}

	function upgradeModern() {
		if (window.ITMSModernCompat) {
			window.ITMSModernCompat.init(document);
		}
	}

	function getVoucherDate() {
		var control = firstField(field("ctlDate")) || byId("ctlDate");
		if (control && typeof control.GetDate === "function") {
			return control.GetDate();
		}
		if (control && typeof control.getDate === "function") {
			return control.getDate();
		}
		if (control && control.Value != null) {
			return control.Value;
		}
		return control && control.value || "";
	}

	function pageName() {
		return String(window.location.pathname || "").split(/[\/\\]/).pop().toLowerCase();
	}

	function targetAction() {
		return /^voudnthrmiscpay\.asp$/i.test(pageName()) ? "VouDNForMiscPayGenerate.asp" : "VouCNForMiscRecGenerate.asp";
	}

	function saveXml() {
		var frm = form();
		setValue("hVouDate", getVoucherDate());
		if (frm) {
			frm.action = targetAction();
			frm.submit();
		}
		return true;
	}

	function enableApproval(source) {
		var approver = firstField(field("selUserId"));
		if (!approver) {
			return;
		}
		if (source && source.value === "Y") {
			approver.disabled = false;
			return;
		}
		approver.selectedIndex = 0;
		approver.disabled = true;
	}

	function xmlRoot(value) {
		var parsed;
		if (!value) {
			return null;
		}
		if (value.documentElement) {
			return value.documentElement;
		}
		if (value._doc && value._doc.documentElement) {
			return value._doc.documentElement;
		}
		if (typeof value === "string" && trim(value).charAt(0) === "<" && window.DOMParser) {
			parsed = new DOMParser().parseFromString(value, "text/xml");
			return parsed && parsed.documentElement || null;
		}
		return null;
	}

	function attr(node, name) {
		var lowerName;
		var i;
		if (!node || !node.attributes) {
			return "";
		}
		if (node.getAttribute && node.getAttribute(name) != null) {
			return node.getAttribute(name);
		}
		lowerName = String(name).toLowerCase();
		for (i = 0; i < node.attributes.length; i += 1) {
			if (String(node.attributes[i].name || "").toLowerCase() === lowerName) {
				return node.attributes[i].value || "";
			}
		}
		return "";
	}

	function childElements(node) {
		var result = [];
		var i;
		if (!node) {
			return result;
		}
		for (i = 0; i < node.childNodes.length; i += 1) {
			if (node.childNodes[i].nodeType === 1) {
				result.push(node.childNodes[i]);
			}
		}
		return result;
	}

	function applyGlReturn(value) {
		var root = xmlRoot(value);
		var nodes = childElements(root);
		var parts;
		if (nodes.length) {
			setValue("hCrAccHead", attr(nodes[0], "RetField0"));
			setHtml("spAccHead", attr(nodes[0], "RetField5"));
			return true;
		}
		if (attr(root, "RetField0")) {
			setValue("hCrAccHead", attr(root, "RetField0"));
			setHtml("spAccHead", attr(root, "RetField5"));
			return true;
		}
		if (typeof value === "string") {
			parts = value.split(":");
			if (parts.length > 1) {
				setValue("hCrAccHead", parts[0]);
				setHtml("spAccHead", parts[5] || "");
				return true;
			}
		}
		return false;
	}

	function openGlDialog(url, callback) {
		var features = "dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No";
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog(url, "", features, callback || function () {});
		} else {
			window.open(url, "_blank", "height=480,width=420,resizable=no,status=no,scrollbars=yes");
		}
	}

	function continueGlHead(url) {
		openGlDialog(url, function (returnedValue) {
			var root = xmlRoot(returnedValue);
			var action = trim(attr(root, "Action")).toUpperCase();
			var query = attr(root, "PassQuery");
			var parts;
			if (action === "CLOSE") {
				return;
			}
			if (action && action !== "DONE" && query) {
				continueGlHead(url.replace(/\?.*$/, "") + "?" + query);
				return;
			}
			if (typeof returnedValue === "string") {
				parts = returnedValue.split(":");
				if (parts.length <= 1 && trim(returnedValue) !== "") {
					continueGlHead(url.replace(/\?.*$/, "") + "?" + returnedValue);
					return;
				}
			}
			applyGlReturn(returnedValue);
		});
	}

	function showGlHead(orgId) {
		var bookNo = valueOf("hBookcode", valueOf("hBookCode"));
		var accHead = valueOf("hCrAccHead", "0");
		var url = "GLHeadSelection.asp?orgId=" + encodeURIComponent(orgId) +
			"&BookId=01&BookNo=" + encodeURIComponent(bookNo) +
			"&AccHead=" + encodeURIComponent(accHead || "0");
		continueGlHead(url);
	}

	function accHead(source) {
		var select = source || firstField(field("SelAccountHd"));
		if (select && select.selectedIndex > 0 && select.value === "G") {
			showGlHead(valueOf("hOrgId", valueOf("hOrgID")));
		}
	}

	function noop() {
		return true;
	}

	function install() {
		upgradeModern();
		window.SaveXML = saveXml;
		window.EnbApp = enableApproval;
		window.AccHead = accHead;
		window.showGLHead = showGlHead;
		window.SetRetVal = noop;
		window.ResetTax = noop;
		window.ReTotalCr = noop;
		ready(upgradeModern);
	}

	window.ITMSMiscThroughNoteCompat = {
		install: install
	};
}(window, document));
