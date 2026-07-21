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

	function field(name) {
		var frm = form();
		return frm && frm.elements ? frm.elements[name] : null;
	}

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function selectedOption(select) {
		return select && select.selectedIndex >= 0 ? select.options[select.selectedIndex] : null;
	}

	function setValue(name, value) {
		if (field(name)) {
			field(name).value = value == null ? "" : String(value);
		}
	}

	function buildQuery(params) {
		var parts = [];
		Object.keys(params).forEach(function (key) {
			if (params[key] !== undefined && params[key] !== null) {
				parts.push(encodeURIComponent(key) + "=" + encodeURIComponent(params[key]));
			}
		});
		return parts.join("&");
	}

	function openDialog(url, args, features, callback) {
		ensureCompat();
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog(url, args, features, callback);
			return;
		}
		if (window.showModalDialog) {
			callback(window.showModalDialog(url, args, features));
			return;
		}
		var popup = window.open(url, "_blank", "width=650,height=500,resizable=yes,scrollbars=yes");
		var timer = window.setInterval(function () {
			if (popup && popup.closed) {
				window.clearInterval(timer);
				callback(popup.returnValue);
			}
		}, 250);
	}

	function parseXml(text) {
		return new DOMParser().parseFromString(text || "<Root/>", "text/xml");
	}

	function elementChildren(node) {
		var result = [];
		var children = node && node.childNodes || [];
		for (var i = 0; i < children.length; i += 1) {
			if (children[i].nodeType === 1) {
				result.push(children[i]);
			}
		}
		return result;
	}

	function getAttr(node, name) {
		return trim(node && node.getAttribute ? node.getAttribute(name) : "");
	}

	function selectedClassPayload(returnValue) {
		var firstPart = trim(returnValue).split("*****")[0] || "";
		return firstPart && firstPart !== "-1" ? firstPart : "";
	}

	function clearItems() {
		if (field("selFrombox")) {
			field("selFrombox").options.length = 0;
		}
		if (field("selTobox")) {
			field("selTobox").options.length = 0;
		}
	}

	function addClassOption(code, name) {
		var select = field("selClass");
		if (!select || !code) {
			return;
		}
		for (var i = 0; i < select.options.length; i += 1) {
			if (select.options[i].value === code) {
				select.selectedIndex = i;
				return;
			}
		}
		select.options[select.options.length] = new Option(name || code, code);
		select.selectedIndex = select.options.length - 1;
	}

	function populateClassDetails(classPayload) {
		var xhr = new XMLHttpRequest();
		xhr.open("GET", "../Master/XMLSelectItemClass.asp?" + buildQuery({
			sOrgID: field("selUnit").value,
			sText: classPayload
		}), false);
		xhr.send(null);
		if (!trim(xhr.responseText)) {
			return;
		}
		var doc = parseXml(xhr.responseText);
		var classes = elementChildren(doc.documentElement);
		for (var i = 0; i < classes.length; i += 1) {
			addClassOption(getAttr(classes[i], "CLASSCODE"), getAttr(classes[i], "CLASSNAME"));
		}
	}

	function AddClass() {
		var unit = field("selUnit");
		var unitOption = selectedOption(unit);
		if (!unit || unit.selectedIndex <= 0) {
			alert("Select Organization");
			if (unit) {
				unit.focus();
			}
			return;
		}
		var url = "/include/ClassificationSelectPop.asp?" + buildQuery({
			sIType: "NO",
			sOrgID: unit.value,
			sITypename: "NO",
			SelMode: "M"
		});
		openDialog(url, "Classification", "dialogHeight:500px;dialogWidth:650px;center:Yes;help:No;resizable:No;status:No", function (result) {
			var classPayload = selectedClassPayload(result || "");
			if (!classPayload) {
				return;
			}
			populateClassDetails(classPayload);
			setValue("hOrgName", unitOption ? unitOption.text : "");
			window.popItmDisplay();
		});
	}

	function popItmDisplay() {
		var select = field("selClass");
		if (!select || select.selectedIndex <= 0) {
			clearItems();
			return;
		}
		var xhr = new XMLHttpRequest();
		xhr.open("GET", "../../Common/XMLGetItemSelectRel.asp?" + buildQuery({
			PageSize: 500,
			hSelectMode: "M",
			hClassCodes: select.value,
			Eligible: "N:N:Y:N",
			Disp: "ALL"
		}), false);
		xhr.send(null);
		clearItems();
		if (!trim(xhr.responseText)) {
			alert("No Item found for Classification selected.");
			return;
		}
		var doc = parseXml(xhr.responseText);
		var items = elementChildren(doc.documentElement);
		for (var i = 0; i < items.length; i += 1) {
			var itemCode = getAttr(items[i], "ItemCode");
			var classCode = getAttr(items[i], "ClassCode");
			var itemName = getAttr(items[i], "ItemName");
			var companyCode = getAttr(items[i], "ComItemCode");
			var text = (companyCode ? companyCode + " / " : "") + itemName;
			field("selFrombox").options[field("selFrombox").options.length] = new Option(text, itemCode + ":" + classCode + ":" + itemName);
		}
	}

	function selectedItems() {
		var select = field("selTobox");
		var items = [];
		var parts;
		for (var i = 0; select && i < select.options.length; i += 1) {
			parts = String(select.options[i].value || "").split(":");
			items.push({
				itemCode: trim(parts[0]),
				classCode: trim(parts[1]),
				itemName: trim(parts.slice(2).join(":")) || select.options[i].text
			});
		}
		return items;
	}

	function CheckSubmit() {
		var unit = field("selUnit");
		var classSelect = field("selClass");
		var classOption = selectedOption(classSelect);
		var control = field("selControl");
		var controlOption = selectedOption(control);
		var items = selectedItems();
		if (!unit || unit.selectedIndex <= 0) {
			alert("Select Organization");
			unit && unit.focus();
			return;
		}
		if (!classSelect || classSelect.selectedIndex <= 0) {
			alert("Select Classification");
			classSelect && classSelect.focus();
			return;
		}
		if (!items.length) {
			alert("Select Item");
			return;
		}
		if (!control || control.selectedIndex <= 0) {
			alert("Select Control");
			control && control.focus();
			return;
		}
		setValue("hSelectedValue", items.map(function (item) { return item.itemCode; }).join("|") + "|");
		setValue("hItemNames", items.map(function (item) { return item.itemName; }).join("|") + "|");
		setValue("hClassName", classOption ? classOption.text : "");
		setValue("hOrgName", selectedOption(unit) ? selectedOption(unit).text : "");
		form().action = controlOption ? controlOption.value : control.value;
		form().submit();
	}

	window.AddClass = AddClass;
	window.popItmDisplay = popItmDisplay;
	window.CheckSubmit = CheckSubmit;
}(window, document));
