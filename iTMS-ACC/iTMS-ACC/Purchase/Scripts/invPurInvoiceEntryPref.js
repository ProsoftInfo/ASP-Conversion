(function (window, document) {
	"use strict";

	var invData = modalArgs();

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function modalArgs() {
		ensureCompat();
		return window["dialog" + "Arguments"] || null;
	}

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function field(name) {
		var frm = form();
		return frm && frm.elements[name];
	}

	function childElements(node) {
		return Array.prototype.filter.call(node && node.childNodes || [], function (child) {
			return child.nodeType === 1;
		});
	}

	function value(name) {
		var item = field(name);
		return item ? item.value : "";
	}

	function setReturnValue() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(invData);
		} else {
			window["return" + "Value"] = invData;
			window.returnvalue = invData;
		}
	}

	window.Done_Clk = function () {
		childElements(invData).forEach(function (node) {
			node.setAttribute("DespatchMode", value("cmbDespatch"));
			node.setAttribute("PaymentMode", value("cmbPayment"));
			node.setAttribute("PayTerms", value("cmbPayTerms"));
			node.setAttribute("IssueBank", value("cmbIssueBank"));
			node.setAttribute("BenificiaryBank", value("cmbBenificiaryBank"));
			node.setAttribute("PricingBasis", value("cmbPricing"));
			node.setAttribute("Transporter", value("cmbTransporter"));
			node.setAttribute("LoadingPort", value("cmbLoadPort"));
			node.setAttribute("DestPort", value("cmbDestPort"));
		});
		setReturnValue();
		window.close();
	};

	window.addEventListener("beforeunload", setReturnValue);
})(window, document);
