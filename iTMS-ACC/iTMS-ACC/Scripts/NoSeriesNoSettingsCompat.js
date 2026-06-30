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
		return frm && (frm.elements[name] || frm[name]) || document.getElementById(name) || null;
	}

	function value(name) {
		var item = field(name);
		return item ? item.value : "";
	}

	function setValue(name, data) {
		var item = field(name);
		if (item) {
			item.value = data == null ? "" : String(data);
		}
	}

	function setText(id, data) {
		var item = document.getElementById(id) || field(id);
		if (item) {
			item.textContent = data == null ? "" : String(data);
		}
	}

	function upgradeXml() {
		if (window.ITMSModernCompat) {
			window.ITMSModernCompat.upgradeXmlIslands(document);
		}
	}

	function xmlObject(name) {
		var element;
		upgradeXml();
		element = document.getElementById(name);
		return window[name] || document[name] || element && element._itmsXmlIsland || element || null;
	}

	function xmlDocument(name) {
		var data = xmlObject(name);
		return data && (data.XMLDocument || data._doc || data) || null;
	}

	function xmlRoot(name) {
		var data = xmlObject(name);
		var doc = xmlDocument(name);
		return data && data.documentElement || doc && doc.documentElement || null;
	}

	function childElements(node) {
		var result = [];
		for (var i = 0; node && i < node.childNodes.length; i += 1) {
			if (node.childNodes[i].nodeType === 1) {
				result.push(node.childNodes[i]);
			}
		}
		return result;
	}

	function clearChildren(node) {
		while (node && node.firstChild) {
			node.removeChild(node.firstChild);
		}
	}

	function serializeXml(name) {
		var doc = xmlDocument(name);
		if (doc && typeof XMLSerializer !== "undefined") {
			return new XMLSerializer().serializeToString(doc);
		}
		return doc && doc.xml || "";
	}

	function loadXml(name, text) {
		var data = xmlObject(name);
		if (data && typeof data.loadXML === "function") {
			data.loadXML(text || "<root/>");
		} else if (data) {
			data._doc = new DOMParser().parseFromString(text || "<root/>", "text/xml");
		}
		return xmlRoot(name);
	}

	function loadXmlFromUrl(name, url) {
		var xhr = new XMLHttpRequest();
		var text;
		xhr.open("GET", url, false);
		xhr.send(null);
		text = xhr.responseText || "";
		if (trim(text) === "" && xhr.responseXML && xhr.responseXML.documentElement) {
			text = new XMLSerializer().serializeToString(xhr.responseXML);
		}
		if (trim(text) !== "") {
			return loadXml(name, text);
		}
		return xmlRoot(name);
	}

	function submitTo(action) {
		var frm = form();
		frm.action = action;
		frm.submit();
	}

	function setSelectByValue(select, data) {
		if (!select) {
			return false;
		}
		for (var i = 0; i < select.options.length; i += 1) {
			if (trim(select.options[i].value) === trim(data)) {
				select.selectedIndex = i;
				return true;
			}
		}
		return false;
	}

	function appendOption(select, text, optionValue) {
		select.options[select.options.length] = new Option(text || "", optionValue || "");
	}

	function parseClassificationReturn(returnData) {
		var parts = String(returnData == null ? "" : returnData).split("*****");
		var codes = [];
		var categories = [];
		var names = [];
		var valueParts;
		var nameParts;
		if (parts[0] === "-1" || trim(parts[0]) === "") {
			return null;
		}
		valueParts = String(parts[0] || "").split("|");
		valueParts.forEach(function (item) {
			var detail = item.split(":");
			if (detail.length > 1) {
				codes.push(detail[detail.length - 1]);
				categories.push(detail[1]);
			} else if (trim(detail[0]) !== "") {
				categories.push(detail[0].substring(3));
			}
		});
		nameParts = String(parts[1] || "").split("|||");
		nameParts.forEach(function (item) {
			var detail = item.split(":");
			var name = detail[detail.length - 1];
			if (trim(name) !== "") {
				names.push(name);
			}
		});
		return {
			classCode: codes.join(","),
			categoryCode: categories.join(","),
			className: names.join(",")
		};
	}

	function applyClassification(result) {
		var category = field("selCategory");
		if (!result) {
			return;
		}
		setText("txtClass", result.className);
		setValue("hClassCode", result.classCode);
		setValue("hCatCode", result.categoryCode);
		if (setSelectByValue(category, result.categoryCode)) {
			category.disabled = true;
		}
	}

	function selectClassifcation() {
		var org = field("selUnit");
		var orgId = org && org.selectedIndex >= 0 ? org.options[org.selectedIndex].value : "";
		var url = "/include/ClassificationSelectPop.asp?sIType=1&sOrgID=" + encodeURIComponent(orgId) + "&sITypename=&SelMode=M";
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog(url, "Classification", "dialogHeight:500px;dialogWidth:650px;center:Yes;help:No;resizable:No;status:No", function (returnData) {
				applyClassification(parseClassificationReturn(returnData));
			});
		} else {
			alert("Popup support is not available.");
		}
		return false;
	}

	function populateData() {
		if (value("selUnit") === "0") {
			alert("Select Unit");
			if (field("selUnit")) {
				field("selUnit").focus();
			}
			return false;
		}
		submitTo("NoSeriesNoSettings.asp");
		return false;
	}

	function populateDetailOne() {
		if (trim(value("SelActivity")) === "") {
			alert("No Details are available for this Activity");
			return false;
		}
		setValue("hSelModule", value("SelModule"));
		setValue("hSelActivity", value("SelActivity"));
		submitTo("NoSeriesNoSettings.asp");
		return false;
	}

	function populateDetail() {
		var moduleValue = value("SelModule");
		if (moduleValue !== "6" && trim(value("SelActivity")) === "") {
			alert("No Details are available for this Activity");
			return false;
		}
		setValue("hSelModule", moduleValue);
		setValue("hSelActivity", value("SelActivity"));
		submitTo("NoSeriesNoSettings.asp");
		return false;
	}

	function populateActivity(selectedActivity) {
		var select = field("SelActivity");
		var root;
		if (value("selUnit") === "0") {
			alert("Select Unit");
			if (field("selUnit")) {
				field("selUnit").focus();
			}
			setValue("SelModule", "0");
			return false;
		}
		if (value("SelModule") === "0") {
			alert("Select Module");
			if (field("SelModule")) {
				field("SelModule").focus();
			}
			return false;
		}
		if (!select) {
			return false;
		}
		root = loadXmlFromUrl("activity", "XMLModuleActivity.asp?sUnitId=" + encodeURIComponent(value("hdOrgUnit")) + "&hApplicationNo=" + encodeURIComponent(value("SelModule")));
		select.options.length = 0;
		childElements(root).forEach(function (node) {
			appendOption(select, node.getAttribute("ActName"), node.getAttribute("ActCode"));
			if (trim(selectedActivity) !== "" && trim(node.getAttribute("ActCode")) === trim(selectedActivity)) {
				select.selectedIndex = select.options.length - 1;
			}
		});
		return true;
	}

	function closeMe() {
		window.location.href = "../index_admin.asp";
	}

	function noSeriesDocRoot() {
		var doc = xmlDocument("NoSeries");
		var root;
		if (!doc || !doc.createElement) {
			return null;
		}
		while (doc.firstChild) {
			doc.removeChild(doc.firstChild);
		}
		root = doc.createElement("Root");
		doc.appendChild(root);
		root.setAttribute("OrgCode", value("selUnit"));
		return root;
	}

	function addSeriesNode(root, index) {
		var doc = xmlDocument("NoSeries");
		var node = doc.createElement("Series");
		node.setAttribute("SeriesNo", value("hdSeriesNoZ" + index));
		node.setAttribute("SeriesCode", value("hdSeriesCodeZ" + index));
		node.setAttribute("Period", value("hdPeriodZ" + index));
		node.setAttribute("Prefix", value("hdPrefixZ" + index));
		node.setAttribute("Suffix", value("hdSuffixZ" + index));
		node.setAttribute("LastPack", value("hdLastPackNumZ" + index));
		node.setAttribute("NewPack", trim(value("txtPackNumZ" + index)) !== "" ? value("txtPackNumZ" + index) : "");
		root.appendChild(node);
	}

	function validatePackNumbers(rowCount) {
		for (var i = 1; i <= rowCount; i += 1) {
			if (trim(value("txtPackNumZ" + i)) !== "" && isNaN(Number(trim(value("txtPackNumZ" + i))))) {
				alert("Specify numeric value for the packing number");
				if (field("txtPackNumZ" + i)) {
					field("txtPackNumZ" + i).select();
				}
				return false;
			}
		}
		return true;
	}

	function updateMe() {
		var rowCount = parseInt(value("hRowCount"), 10) || 0;
		var root = noSeriesDocRoot();
		var xhr;
		if (!validatePackNumbers(rowCount)) {
			return false;
		}
		for (var i = 1; i <= rowCount; i += 1) {
			addSeriesNode(root, i);
		}
		xhr = new XMLHttpRequest();
		xhr.open("POST", "XMLSave.asp?Name=NoSeriesSettings", false);
		xhr.setRequestHeader("Content-Type", "text/xml");
		xhr.send(serializeXml("NoSeries"));
		if (trim(xhr.responseText) !== "") {
			alert(xhr.responseText);
		} else {
			submitTo("NoSeriesNoSettingsUpdate.asp?Module=" + encodeURIComponent(value("hSelModule")) + "&Activity=" + encodeURIComponent(value("hSelActivity")) + "&ClassCode=" + encodeURIComponent(value("hClassCode")) + "&CatCode=" + encodeURIComponent(value("hCatCode")));
		}
		return false;
	}

	function selUnit() {
		setValue("selUnit", value("hdOrgUnit"));
		setValue("SelModule", value("hSelModule"));
		if (value("SelModule") !== "") {
			populateActivity(trim(value("hSelActivity")) !== "" ? value("hSelActivity") : "0");
			setValue("SelActivity", value("hSelActivity"));
		} else {
			setValue("SelModule", "0");
		}
		setSelectByValue(field("selCategory"), value("hCatCode"));
		setText("txtClass", trim(value("hClassName")) !== "" ? value("hClassName") : value("hCatName"));
	}

	function install() {
		upgradeXml();
		window.SubmitItem = function () {};
		window.SelectClassifcation = selectClassifcation;
		window.PopulateData = populateData;
		window.PopulateDetailOne = populateDetailOne;
		window.PopulateDetail = populateDetail;
		window.PopulateActivity = populateActivity;
		window.closeMe = closeMe;
		window.updateMe = updateMe;
		window.selUnit = selUnit;
	}

	window.ITMSNoSeriesNoSettingsCompat = {
		install: install
	};
}(window, document));
