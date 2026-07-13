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
		return frm && frm.elements ? frm.elements[name] || document.getElementsByName(name)[0] || null : document.getElementsByName(name)[0] || null;
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

	function xmlDoc(name) {
		var object = xmlObject(name);
		var root = xmlRoot(name);
		return object && object.XMLDocument || object && object._doc || root && root.ownerDocument || null;
	}

	function clearChildren(node) {
		while (node && node.firstChild) {
			node.removeChild(node.firstChild);
		}
	}

	function selectedText(select) {
		if (!select || select.selectedIndex < 0 || !select.options[select.selectedIndex]) {
			return "";
		}
		return select.options[select.selectedIndex].text;
	}

	function selectedValues(select) {
		var values = [];
		if (!select) {
			return values;
		}
		for (var i = 0; i < select.options.length; i += 1) {
			if (select.options[i].selected && select.options[i].value !== "select") {
				values.push(select.options[i].value);
			}
		}
		return values;
	}

	function selectedTexts(select) {
		var values = [];
		if (!select) {
			return values;
		}
		for (var i = 0; i < select.options.length; i += 1) {
			if (select.options[i].selected && select.options[i].value !== "select") {
				values.push(select.options[i].text);
			}
		}
		return values;
	}

	function appendElement(doc, parent, name, attrs) {
		var node = doc.createElement(name);
		Object.keys(attrs || {}).forEach(function (key) {
			node.setAttribute(key, attrs[key] == null ? "" : String(attrs[key]));
		});
		parent.appendChild(node);
		return node;
	}

	function returnAndClose(value) {
		if (window.ITMSModalReturnCompat && window.ITMSModalReturnCompat.returnAndClose) {
			window.ITMSModalReturnCompat.returnAndClose(value);
			return;
		}
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(value);
		}
		window.close();
	}

	function Init(itemType) {
		var auto = field("radSort") && field("radSort").length ? field("radSort")[0] : null;
		if (itemType === "FAB" && auto && !field("radSort")[1].checked) {
			auto.checked = true;
			LoadMaxData();
		}
	}

	function maxDataValue(code) {
		var root = xmlRoot("MaxData");
		var nodes = root && root.selectNodes ? root.selectNodes("//DETAILS") : [];
		for (var i = 0; i < nodes.length; i += 1) {
			if (trim(nodes[i].getAttribute("CODE")) === trim(code)) {
				return nodes[i].getAttribute("VALUE") || "";
			}
		}
		return "";
	}

	function LoadMaxData() {
		var first = field("sel1");
		var autoText = field("txtAutomatic");
		if (!first || !autoText) {
			return false;
		}
		if (first.value === "select") {
			autoText.value = "";
			return false;
		}
		autoText.value = maxDataValue(first.value);
		return false;
	}

	function selectedRadioValue(name) {
		var group = field(name);
		if (!group) {
			return "";
		}
		if (group.length === undefined) {
			return group.checked ? group.value : "";
		}
		for (var i = 0; i < group.length; i += 1) {
			if (group[i].checked) {
				return group[i].value;
			}
		}
		return "";
	}

	function codePart(index, itemType) {
		var select = field("sel" + index);
		var mode;
		if (itemType === "FAB" && index === 2 && field("txtAutomatic")) {
			mode = selectedRadioValue("radSort") || "A";
			if (mode === "M") {
				return trim(valueOf("txtManual"));
			}
			if (!valueOf("txtAutomatic")) {
				LoadMaxData();
			}
			return trim(valueOf("txtAutomatic"));
		}
		if (!select || select.value === "select") {
			return "";
		}
		return trim(select.value);
	}

	function buildCode(itemType) {
		var count = Number(valueOf("hCount"));
		var parts = [];
		for (var i = 1; i <= count + 1; i += 1) {
			parts.push(codePart(i, itemType));
		}
		return parts.join("");
	}

	function validateCodeParts(itemType) {
		var count = Number(valueOf("hCount"));
		var select;
		var part;
		for (var i = 1; i <= count + 1; i += 1) {
			select = field("sel" + i);
			part = codePart(i, itemType);
			if (!part) {
				if (itemType === "FAB" && i === 2 && field("txtAutomatic")) {
					alert("Enter or select the running number");
					(field("txtManual") || field("txtAutomatic")).focus();
					return false;
				}
				alert("Select code value");
				if (select && select.focus) {
					select.focus();
				}
				return false;
			}
		}
		return true;
	}

	function CheckData(itemType) {
		var product = trim(valueOf("txtProductname"));
		var weave = trim(valueOf("txtWeave"));
		var warp = selectedTexts(field("selWarp")).join(",");
		var weft = selectedTexts(field("selWeft")).join(",");
		if (itemType !== "FAB") {
			return true;
		}
		if (!product) {
			product = [weave, trim(valueOf("txtWidth")), trim(valueOf("txtWeight")), warp, weft].filter(Boolean).join(" ");
			setValue("txtProductname", product);
		}
		if (field("sel1") && field("sel1").value !== "select") {
			LoadMaxData();
		}
		return true;
	}

	function updateXml(itemType, code) {
		var doc = xmlDoc("OutData");
		var root = xmlRoot("OutData");
		var count = Number(valueOf("hCount"));
		var part;
		if (!doc || !root) {
			return;
		}
		clearChildren(root);
		root.setAttribute("TYPE", itemType || "");
		root.setAttribute("CODE", code || "");
		root.setAttribute("DESCRIPTION", trim(valueOf("txtProductname")));
		for (var i = 1; i <= count + 1; i += 1) {
			part = codePart(i, itemType);
			appendElement(doc, root, "CodePart", {
				INDEX: i,
				VALUE: part,
				TEXT: field("sel" + i) ? selectedText(field("sel" + i)) : part
			});
		}
		if (itemType === "FAB") {
			appendElement(doc, root, "Fabric", {
				WEAVE: valueOf("txtWeave"),
				DENT: valueOf("txtDent"),
				WIDTH: valueOf("txtWidth"),
				ENDS: valueOf("txtEnds"),
				REEDCOUNT: valueOf("txtReedCount"),
				ENDSINCH: valueOf("txtEndsInch"),
				REEDSPACE: valueOf("txtReedSpace"),
				PICKSINCH: valueOf("txtPicksInch"),
				WEIGHT: valueOf("txtWeight"),
				AVGWRAP: valueOf("txtAvgWrap"),
				VARIETY: valueOf("txtVariety"),
				TAPELNE: valueOf("txtTapeLne"),
				WARP: selectedValues(field("selWarp")).join(","),
				WEFT: selectedValues(field("selWeft")).join(",")
			});
		}
	}

	function CheckSubmit(itemType) {
		var product = trim(valueOf("txtProductname"));
		var code;
		if (!product) {
			alert("Enter Product Name");
			field("txtProductname").focus();
			return false;
		}
		if (!validateCodeParts(itemType)) {
			return false;
		}
		if (itemType === "FAB") {
			CheckData(itemType);
		}
		code = buildCode(itemType);
		updateXml(itemType, code);
		returnAndClose(product + "``" + code);
		return false;
	}

	window.Init = Init;
	window.LoadMaxData = LoadMaxData;
	window.CheckData = CheckData;
	window.CheckSubmit = CheckSubmit;
}(window, document));
