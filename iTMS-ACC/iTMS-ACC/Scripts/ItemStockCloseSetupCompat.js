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

	function ensureRoot(name) {
		var doc = xmlDocument(name);
		var root;
		if (!doc || !doc.createElement) {
			return null;
		}
		if (doc.documentElement) {
			return doc.documentElement;
		}
		root = doc.createElement("Root");
		doc.appendChild(root);
		return root;
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
			data.loadXML(text || "<Root/>");
		} else if (data) {
			data._doc = new DOMParser().parseFromString(text || "<Root/>", "text/xml");
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

	function selectedOption(select) {
		if (!select || select.selectedIndex < 0) {
			return null;
		}
		return select.options[select.selectedIndex] || null;
	}

	function selectedOptions(select) {
		var result = [];
		for (var i = 0; select && i < select.options.length; i += 1) {
			if (select.options[i].selected) {
				result.push(select.options[i]);
			}
		}
		return result;
	}

	function appendOption(select, text, optionValue) {
		if (select) {
			select.options[select.options.length] = new Option(text || "", optionValue || "");
		}
	}

	function findChild(parent, nodeName, code) {
		var nodes = childElements(parent);
		for (var i = 0; i < nodes.length; i += 1) {
			if (nodes[i].nodeName === nodeName && trim(nodes[i].getAttribute("Code")) === trim(code)) {
				return nodes[i];
			}
		}
		return null;
	}

	function createNode(name, attrs) {
		var doc = xmlDocument("OutDataXML");
		var node = doc && doc.createElement(name);
		var key;
		if (!node) {
			return null;
		}
		for (key in attrs) {
			if (Object.prototype.hasOwnProperty.call(attrs, key)) {
				node.setAttribute(key, attrs[key] == null ? "" : String(attrs[key]));
			}
		}
		return node;
	}

	function CheckSetup() {
		var frm = form();
		if (!frm) {
			return false;
		}
		if (frm.chkCategory && frm.selCategory) {
			frm.selCategory.disabled = !frm.chkCategory.checked;
		}
		if (frm.chkSubCategory && frm.selSubCategory) {
			frm.selSubCategory.disabled = !frm.chkSubCategory.checked;
		}
		if (frm.chkClassification && frm.selClassification) {
			if (frm.chkClassification.checked) {
				if (frm.selSubCategory) {
					frm.selSubCategory.multiple = false;
					if (frm.selSubCategory.options.length > 0) {
						frm.selSubCategory.selectedIndex = 0;
					}
				}
				populateClassification();
				frm.selClassification.disabled = false;
			} else {
				frm.selClassification.disabled = true;
			}
		}
		return true;
	}

	function populateSubCategory() {
		var frm = form();
		var category = selectedOption(field("selCategory"));
		var root;
		if (!frm) {
			return false;
		}
		if (frm.chkSubCategory) {
			frm.chkSubCategory.checked = false;
		}
		if (frm.chkClassification) {
			frm.chkClassification.checked = false;
		}
		CheckSetup();
		root = loadXmlFromUrl("SubCategory", "XmlGetClassification.asp?Level=0&Category=" + encodeURIComponent(category ? category.value : ""));
		if (childElements(root).length > 0 && frm.selSubCategory) {
			frm.selSubCategory.options.length = 0;
			childElements(root).forEach(function (node) {
				if (node.nodeName === "Level") {
					appendOption(frm.selSubCategory, node.getAttribute("GroupName"), node.getAttribute("GroupCode"));
				}
			});
		}
		return true;
	}

	function populateClassification() {
		var select = field("selSubCategory");
		var subCategory = selectedOption(select);
		var classification = field("selClassification");
		var root = loadXmlFromUrl("Classification", "XmlGetClassification.asp?Level=1&SubCategory=" + encodeURIComponent(subCategory ? subCategory.value : ""));
		if (childElements(root).length > 0 && classification) {
			classification.options.length = 0;
			childElements(root).forEach(function (node) {
				if (node.nodeName === "Level") {
					appendOption(classification, node.getAttribute("GroupName"), node.getAttribute("GroupCode"));
				}
			});
		}
		return true;
	}

	function selectedCategoryInfo() {
		var option = selectedOption(field("selCategory"));
		return {
			code: option ? option.value : "",
			name: option ? option.text : ""
		};
	}

	function selectedSubCategoryInfo() {
		var option = selectedOption(field("selSubCategory"));
		return {
			code: option ? option.value : "",
			name: option ? option.text : ""
		};
	}

	function AddCategorySetup() {
		var frm = form();
		var root = ensureRoot("OutDataXML");
		var categoryInfo = frm && frm.chkCategory && frm.chkCategory.checked ? selectedCategoryInfo() : { code: "", name: "" };
		var subInfo = frm && frm.chkSubCategory && frm.chkSubCategory.checked ? selectedSubCategoryInfo() : { code: "", name: "" };
		var category = findChild(root, "Category", categoryInfo.code);
		var subCategory;
		var options;
		if (!root || !frm) {
			return false;
		}
		if (!category) {
			category = createNode("Category", {
				Code: categoryInfo.code,
				Name: categoryInfo.name,
				Consider: "0"
			});
			root.appendChild(category);
		}
		subCategory = findChild(category, "SubCategory", subInfo.code);
		if (!subCategory && frm.chkSubCategory && frm.chkSubCategory.checked) {
			subCategory = createNode("SubCategory", {
				Code: subInfo.code,
				Name: subInfo.name,
				Consider: "0"
			});
			category.appendChild(subCategory);
		}
		if (frm.chkClassification && frm.chkClassification.checked) {
			options = selectedOptions(frm.selClassification);
			options.forEach(function (option) {
				var classification;
				if (!subCategory) {
					return;
				}
				classification = findChild(subCategory, "Classification", option.value);
				if (!classification) {
					classification = createNode("Classification", {
						Code: option.value,
						Name: option.text,
						Consider: "1"
					});
					subCategory.appendChild(classification);
				}
			});
		} else if (frm.chkSubCategory && frm.chkSubCategory.checked) {
			selectedOptions(frm.selSubCategory).forEach(function (option) {
				var current = findChild(category, "SubCategory", option.value);
				if (current) {
					current.setAttribute("Consider", "1");
				} else {
					current = createNode("SubCategory", {
						Code: option.value,
						Name: option.value,
						Consider: "1"
					});
					category.appendChild(current);
				}
			});
		}
		if (!frm.chkSubCategory || !frm.chkSubCategory.checked) {
			category.setAttribute("Consider", "1");
		} else if (!frm.chkClassification || !frm.chkClassification.checked) {
			if (subCategory) {
				subCategory.setAttribute("Consider", "1");
			}
		}
		DisplayTable();
		return true;
	}

	function ClearTable() {
		var grid = document.getElementById("tblSetup");
		while (grid && grid.rows.length > 1) {
			grid.deleteRow(1);
		}
	}

	function addTextCell(row, className, align, text) {
		var cell = row.insertCell();
		cell.className = className || "";
		if (align) {
			cell.align = align;
		}
		cell.textContent = text == null ? "" : String(text);
		return cell;
	}

	function addSetupRow(serial, checkValue, categoryName, subCategoryName, classificationName) {
		var grid = document.getElementById("tblSetup");
		var row = grid && grid.insertRow(grid.rows.length);
		var cell;
		var input;
		if (!row) {
			return;
		}
		addTextCell(row, "ExcelSerial", "Center", serial);
		cell = row.insertCell();
		cell.className = "ExcelDisplayCell";
		cell.align = "Center";
		input = document.createElement("input");
		input.type = "checkbox";
		input.name = "ChkSetupZ" + serial;
		input.value = checkValue;
		input.size = 11;
		input.className = "Formelem";
		cell.appendChild(input);
		addTextCell(row, "ExcelDisplayCell", "Left", categoryName);
		addTextCell(row, "ExcelDisplayCell", "Left", subCategoryName);
		addTextCell(row, "ExcelDisplayCell", "Left", classificationName);
	}

	function DisplayTable() {
		var root = ensureRoot("OutDataXML");
		var serial = 0;
		ClearTable();
		childElements(root).forEach(function (category) {
			var categoryCode;
			var categoryName;
			if (category.nodeName !== "Category") {
				return;
			}
			categoryCode = category.getAttribute("Code") || "";
			categoryName = category.getAttribute("Name") || "";
			if (category.getAttribute("Consider") === "1") {
				serial += 1;
				addSetupRow(serial, categoryCode + "::", categoryName, "N/A", "N/A");
				return;
			}
			childElements(category).forEach(function (subCategory) {
				var subCategoryCode;
				var subCategoryName;
				if (subCategory.nodeName !== "SubCategory") {
					return;
				}
				subCategoryCode = subCategory.getAttribute("Code") || "";
				subCategoryName = subCategory.getAttribute("Name") || "";
				if (subCategory.getAttribute("Consider") === "1") {
					serial += 1;
					addSetupRow(serial, categoryCode + ":" + subCategoryCode + ":", categoryName, subCategoryName, "N/A");
					return;
				}
				childElements(subCategory).forEach(function (classification) {
					if (classification.nodeName === "Classification" && classification.getAttribute("Consider") === "1") {
						serial += 1;
						addSetupRow(serial, categoryCode + ":" + subCategoryCode + ":" + (classification.getAttribute("Code") || ""), categoryName, subCategoryName, classification.getAttribute("Name") || "");
					}
				});
			});
		});
		setValue("hRow", serial);
		return true;
	}

	function LoadXML() {
		loadXmlFromUrl("OutDataXML", "XMLGetCategorySetup.asp");
		return true;
	}

	function checkedSetupValues() {
		var count = parseInt(value("hRow"), 10) || 0;
		var selected = [];
		var item;
		for (var i = 1; i <= count; i += 1) {
			item = field("ChkSetupZ" + i);
			if (item && item.checked) {
				selected.push(item.value);
			}
		}
		return selected;
	}

	function DeleteSetUp() {
		var selected = checkedSetupValues();
		var parts;
		var root;
		if (selected.length !== 1) {
			alert("Select single entry to delete");
			return false;
		}
		parts = selected[0].split(":");
		root = ensureRoot("OutDataXML");
		childElements(root).forEach(function (category) {
			if (category.nodeName !== "Category" || trim(category.getAttribute("Code")) !== trim(parts[0])) {
				return;
			}
			if (category.getAttribute("Consider") === "1") {
				category.setAttribute("Consider", "0");
				return;
			}
			childElements(category).forEach(function (subCategory) {
				if (subCategory.nodeName !== "SubCategory" || trim(subCategory.getAttribute("Code")) !== trim(parts[1])) {
					return;
				}
				if (subCategory.getAttribute("Consider") === "1") {
					subCategory.setAttribute("Consider", "0");
					return;
				}
				childElements(subCategory).forEach(function (classification) {
					if (classification.nodeName === "Classification" && trim(classification.getAttribute("Code")) === trim(parts[2]) && classification.getAttribute("Consider") === "1") {
						classification.setAttribute("Consider", "0");
					}
				});
			});
		});
		DisplayTable();
		return true;
	}

	function SetupCategory() {
		var xhr = new XMLHttpRequest();
		xhr.open("POST", "ItemStockCloseSetupInsert.asp", false);
		xhr.setRequestHeader("Content-Type", "text/xml");
		xhr.send(serializeXml("OutDataXML"));
		if (trim(xhr.responseText) !== "") {
			alert(xhr.responseText);
		} else {
			alert("Records updated successfully");
			LoadXML();
			DisplayTable();
		}
		return false;
	}

	function install() {
		upgradeXml();
		window.CheckSetup = CheckSetup;
		window.populateSubCategory = populateSubCategory;
		window.populateClassification = populateClassification;
		window.AddCategorySetup = AddCategorySetup;
		window.ClearTable = ClearTable;
		window.DisplayTable = DisplayTable;
		window.LoadXML = LoadXML;
		window.DeleteSetUp = DeleteSetUp;
		window.SetupCategory = SetupCategory;
	}

	window.ITMSItemStockCloseSetupCompat = {
		install: install
	};
}(window, document));
