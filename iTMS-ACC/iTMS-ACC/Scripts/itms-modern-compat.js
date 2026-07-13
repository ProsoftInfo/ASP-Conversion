(function () {
	"use strict";

	var DATE_PICKER_CLSID = "01e5bf20-f919-44e6-a698-cf7fd7c7d6cd";
	var TREE_CLSIDS = {
		account: "355ceafa-cb06-4345-8384-d0725c8a3048",
		classification: "c93b7cfc-55f8-49ba-bef7-b6cddfe2eb10",
		itemClassification: "39b53116-8621-41ea-afb9-3d15df15c41e",
		itemClassificationMulti: "ef0b79dd-fbe8-49ce-bed8-f7d04a9b7447"
	};
	var treeControls = {};

	function onReady(fn) {
		if (document.readyState === "loading") {
			document.addEventListener("DOMContentLoaded", fn);
		} else {
			fn();
		}
	}

	function hasClassId(el, clsid) {
		return String(el.getAttribute("classid") || "").toLowerCase().indexOf(clsid) !== -1;
	}

	function treeKindFromObject(el) {
		if (!el || !el.getAttribute) {
			return "";
		}
		if (el.hasAttribute("data-tree-kind")) {
			return el.getAttribute("data-tree-kind") || "";
		}
		if (hasClassId(el, TREE_CLSIDS.classification)) {
			return "classification";
		}
		if (hasClassId(el, TREE_CLSIDS.itemClassification) || hasClassId(el, TREE_CLSIDS.itemClassificationMulti)) {
			return "item-classification";
		}
		if (hasClassId(el, TREE_CLSIDS.account)) {
			return "account";
		}
		return "";
	}

	function isTreeObject(el) {
		return !!treeKindFromObject(el);
	}

	function readParams(objectEl) {
		var params = {};
		var nodes = objectEl.getElementsByTagName("param");
		for (var i = 0; i < nodes.length; i += 1) {
			var name = nodes[i].getAttribute("name");
			if (name) {
				params[name.toLowerCase()] = nodes[i].getAttribute("value") || "";
			}
		}
		[
			["dsn", "data-dsn"],
			["listname", "data-list-name"],
			["groupvalue", "data-group-value"],
			["groupname", "data-group-name"],
			["headvalue", "data-head-value"],
			["headname", "data-head-name"],
			["treekind", "data-tree-kind"],
			["itype", "data-itype"],
			["classification", "data-classification"],
			["classificationname", "data-classification-name"]
		].forEach(function (entry) {
			if (objectEl.hasAttribute(entry[1])) {
				params[entry[0]] = objectEl.getAttribute(entry[1]) || "";
			}
		});
		return params;
	}

	function pad2(value) {
		return value < 10 ? "0" + value : String(value);
	}

	function parseLegacyDate(value) {
		var match;
		var text;
		if (value instanceof Date && !isNaN(value.getTime())) {
			return new Date(value.getFullYear(), value.getMonth(), value.getDate());
		}
		if (value == null) {
			return null;
		}
		text = String(value).replace(/\u00a0/g, " ").trim();
		if (!text) {
			return null;
		}
		match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
		if (match) {
			return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
		}
		match = text.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/);
		if (match) {
			var year = Number(match[3]);
			if (year < 100) {
				year += 2000;
			}
			return new Date(year, Number(match[2]) - 1, Number(match[1]));
		}
		var parsed = new Date(text);
		if (!isNaN(parsed.getTime())) {
			return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
		}
		return null;
	}

	function toIsoDate(value) {
		var date = parseLegacyDate(value);
		if (!date) {
			return "";
		}
		return date.getFullYear() + "-" + pad2(date.getMonth() + 1) + "-" + pad2(date.getDate());
	}

	function toDisplayDate(value) {
		var date = parseLegacyDate(value);
		if (!date) {
			return "";
		}
		return pad2(date.getDate()) + "/" + pad2(date.getMonth() + 1) + "/" + date.getFullYear();
	}

	function setInputDate(input, value) {
		var isoValue = toIsoDate(value);
		input.value = isoValue;
		syncRelatedHiddenFields(input);
		return isoValue;
	}

	function getDateValue(input) {
		return toDisplayDate(input.value);
	}

	function candidateHiddenFields(input) {
		var id = String(input.id || input.name || "").toLowerCase();
		var candidates = [];
		if (id.indexOf("from") !== -1 || id.indexOf("frm") !== -1) {
			candidates.push("hFromDate", "hFrmDate");
		}
		if (id.indexOf("to") !== -1) {
			candidates.push("hToDate");
		}
		if (id.indexOf("till") !== -1) {
			candidates.push("hTillDate", "hToDate");
		}
		if (id.indexOf("invoice") !== -1 || id === "ctldate") {
			candidates.push("hInvDate", "hInvoiceDate", "hVouDate", "hDate");
		}
		if (id.indexOf("cheque") !== -1) {
			candidates.push("hChequeDate", "hChqDate");
		}
		if (id.indexOf("issue") !== -1) {
			candidates.push("hIssueDate");
		}
		if (id.indexOf("acc") !== -1) {
			candidates.push("hAccDate");
		}
		return candidates;
	}

	function getFormElement(form, name) {
		if (!form || !name) {
			return null;
		}
		return form.elements[name] || form.querySelector('[name="' + name.replace(/"/g, '\\"') + '"]');
	}

	function syncRelatedHiddenFields(input) {
		var form = input.form;
		var value = getDateValue(input);
		var candidates = candidateHiddenFields(input);
		for (var i = 0; i < candidates.length; i += 1) {
			var field = getFormElement(form, candidates[i]);
			if (field && field.type === "hidden") {
				field.value = value;
			}
		}
	}

	function decorateDateInput(input) {
		if (!input || input.getAttribute("data-itms-date-ready") === "1") {
			return input;
		}
		input.setAttribute("data-itms-date-ready", "1");
		if (!/\bitms-date-picker\b/.test(input.className)) {
			input.className = (input.className ? input.className + " " : "") + "itms-date-picker";
		}
		if (!input.type || input.type.toLowerCase() !== "date") {
			try {
				input.type = "date";
			} catch (ignore) {}
		}

		input.GetDate = function () {
			return getDateValue(input);
		};
		input.getDate = input.GetDate;
		input.SetDate = function (value) {
			setInputDate(input, value);
		};
		input.setDate = input.SetDate;
		input.SetMinDate = function (value) {
			input.min = toIsoDate(value);
		};
		input.setMinDate = input.SetMinDate;
		input.SetMaxDate = function (value) {
			input.max = toIsoDate(value);
		};
		input.setMaxDate = input.SetMaxDate;

		try {
			Object.defineProperty(input, "Value", {
				get: function () {
					return getDateValue(input);
				},
				set: function (value) {
					setInputDate(input, value);
				}
			});
			Object.defineProperty(input, "Text", {
				get: function () {
					return getDateValue(input);
				},
				set: function (value) {
					setInputDate(input, value);
				}
			});
			Object.defineProperty(input, "Day", {
				get: function () {
					var date = parseLegacyDate(input.value);
					return date ? date.getDate() : 0;
				}
			});
			Object.defineProperty(input, "Month", {
				get: function () {
					var date = parseLegacyDate(input.value);
					return date ? date.getMonth() + 1 : 0;
				}
			});
			Object.defineProperty(input, "Year", {
				get: function () {
					var date = parseLegacyDate(input.value);
					return date ? date.getFullYear() : 0;
				}
			});
		} catch (ignoreProperties) {}

		input.addEventListener("change", function () {
			syncRelatedHiddenFields(input);
		});
		input.addEventListener("blur", function () {
			syncRelatedHiddenFields(input);
		});

		if (!input.value) {
			var candidates = candidateHiddenFields(input);
			for (var i = 0; i < candidates.length; i += 1) {
				var field = getFormElement(input.form, candidates[i]);
				if (field && field.value) {
					setInputDate(input, field.value);
					break;
				}
			}
		}
		if (!input.value) {
			var currentDateField = getFormElement(input.form, "hCurrDate");
			if (currentDateField && currentDateField.value) {
				setInputDate(input, currentDateField.value);
			}
		}
		if (!input.value) {
			setInputDate(input, new Date());
		}
		return input;
	}

	function createDateInputFromObject(objectEl) {
		var input = document.createElement("input");
		var id = objectEl.id || objectEl.getAttribute("name") || "itmsDate" + Math.floor(Math.random() * 1000000);
		input.type = "date";
		input.id = id;
		input.name = id;
		input.className = objectEl.className || objectEl.getAttribute("class") || "FormElem";
		input.title = objectEl.title || "";
		if (objectEl.getAttribute("onblur")) {
			input.setAttribute("onblur", objectEl.getAttribute("onblur"));
		}
		if (objectEl.getAttribute("width")) {
			input.style.width = objectEl.getAttribute("width");
		}
		objectEl.parentNode.replaceChild(input, objectEl);
		decorateDateInput(input);
		return input;
	}

	function upgradeDatePickers(root) {
		var scope = root || document;
		var objects = scope.querySelectorAll("object[classid]");
		for (var i = objects.length - 1; i >= 0; i -= 1) {
			if (hasClassId(objects[i], DATE_PICKER_CLSID)) {
				createDateInputFromObject(objects[i]);
			}
		}
		var inputs = scope.querySelectorAll("input[data-itms-datepicker], input.itms-date-picker");
		for (var j = 0; j < inputs.length; j += 1) {
			decorateDateInput(inputs[j]);
		}
	}

	function afterColon(value) {
		var text = String(value || "");
		var index = text.indexOf(":");
		return index === -1 ? text : text.substring(index + 1);
	}

	function defineControlProperty(control, state, name, key) {
		try {
			Object.defineProperty(control, name, {
				get: function () {
					return state[key];
				},
				set: function (value) {
					state[key] = value == null ? "" : String(value);
					if (key === "dsn") {
						refreshTree(control);
					}
				}
			});
		} catch (ignore) {}
	}

	function createTreeControlFromObject(objectEl) {
		var params = readParams(objectEl);
		var id = objectEl.id || objectEl.getAttribute("name") || "itmsTree" + Math.floor(Math.random() * 1000000);
		var wrapper = document.createElement("span");
		var control = document.createElement("input");
		var host = document.createElement("div");
		var state = {
			dsn: params.dsn || "",
			listName: params.listname || "",
			groupValue: params.groupvalue || "0",
			groupName: params.groupname || "",
			headValue: params.headvalue || "0",
			headName: params.headname || "",
			treeKind: params.treekind || treeKindFromObject(objectEl),
			iType: params.itype || "NO:NO:NO",
			selectedKey: "",
			selectedText: "",
			selectedFullPath: "",
			classification: params.classification || "",
			classificationName: params.classificationname || ""
		};

		wrapper.className = "itms-tree-wrapper";
		control.type = "hidden";
		control.id = id;
		control.name = id;
		control.value = state.groupValue;
		host.id = id + "_tree";
		host.className = "itms-tree";
		host.tabIndex = 0;
		host.setAttribute("role", "tree");
		host.style.width = objectEl.getAttribute("data-width") || objectEl.getAttribute("width") || objectEl.style.width || "263px";
		host.style.height = objectEl.getAttribute("data-height") || objectEl.getAttribute("height") || objectEl.style.height || "353px";
		host.innerHTML = '<div class="itms-tree-status">Loading...</div>';

		wrapper.appendChild(control);
		wrapper.appendChild(host);
		objectEl.parentNode.replaceChild(wrapper, objectEl);

		treeControls[id] = { control: control, host: host, state: state, groups: {} };
		defineControlProperty(control, state, "DSN", "dsn");
		defineControlProperty(control, state, "ListName", "listName");
		defineControlProperty(control, state, "GroupValue", "groupValue");
		defineControlProperty(control, state, "GroupName", "groupName");
		defineControlProperty(control, state, "HeadValue", "headValue");
		defineControlProperty(control, state, "HeadName", "headName");
		defineControlProperty(control, state, "classification", "classification");
		defineControlProperty(control, state, "Classification", "classification");
		defineControlProperty(control, state, "classificationName", "classificationName");
		defineControlProperty(control, state, "ClassificationName", "classificationName");
		defineTreeValueProperty(control, "GetKey", state, "selectedKey");
		defineTreeValueProperty(control, "GetText", state, "selectedText");
		defineTreeValueProperty(control, "GetFullPath", state, "selectedFullPath");
		defineTreeITypeProperty(control, state);
		control.Refresh = function () {
			refreshTree(control);
		};
		control.refresh = control.Refresh;
		control.populateTree = control.Refresh;
		control.PopulateTree = control.Refresh;
		try {
			window[id] = control;
			document[id] = control;
		} catch (ignoreControlAlias) {}
		refreshTree(control);
		return control;
	}

	function defineTreeValueProperty(control, name, state, key) {
		try {
			Object.defineProperty(control, name, {
				get: function () {
					return state[key] || "";
				},
				set: function (value) {
					state[key] = value == null ? "" : String(value);
				}
			});
		} catch (ignore) {}
	}

	function defineTreeITypeProperty(control, state) {
		try {
			Object.defineProperty(control, "IType", {
				get: function () {
					return state.iType || "";
				},
				set: function (value) {
					state.iType = value == null ? "" : String(value);
					refreshTree(control);
				}
			});
		} catch (ignore) {}
	}

	function readAttributeAny(element, names) {
		for (var i = 0; i < names.length; i += 1) {
			if (element.getAttribute(names[i]) != null) {
				return element.getAttribute(names[i]) || "";
			}
		}
		return "";
	}

	function parseTreeXml(xmlDoc) {
		var groups = [];
		var heads = [];
		var elements = xmlDoc.getElementsByTagName("*");
		for (var i = 0; i < elements.length; i += 1) {
			var tagName = elements[i].tagName || "";
			if (/group$/i.test(tagName) && (elements[i].getAttribute("GroupCode") != null || elements[i].getAttribute("Code") != null || elements[i].getAttribute("Key") != null)) {
				var code = readAttributeAny(elements[i], ["GroupCode", "Key", "Code"]);
				var name = readAttributeAny(elements[i], ["GroupName", "Text", "Name"]);
				var parent = readAttributeAny(elements[i], ["ParentCode", "ParentGroup"]);
				if (!parent && elements[i].getAttribute("Key") != null) {
					parent = readAttributeAny(elements[i], ["Code"]);
				}
				if (parent === code || parent === "GRP") {
					parent = "";
				}
				groups.push({
					code: code || "0",
					name: name || code || "",
					parent: parent || ""
				});
			}
			if (/head$/i.test(tagName) && elements[i].getAttribute("HeadCode") != null) {
				heads.push({
					code: elements[i].getAttribute("HeadCode") || "0",
					value: afterColon(elements[i].getAttribute("HeadCode") || "0"),
					name: elements[i].getAttribute("HeadName") || "",
					parent: elements[i].getAttribute("ParentCode") || ""
				});
			}
		}
		return { groups: groups, heads: heads };
	}

	function buildTreeModel(parsed) {
		var map = {};
		var roots = [];
		for (var i = 0; i < parsed.groups.length; i += 1) {
			var group = parsed.groups[i];
			group.children = [];
			group.heads = [];
			map[group.code] = group;
		}
		for (var j = 0; j < parsed.groups.length; j += 1) {
			var current = parsed.groups[j];
			if (current.parent && current.parent !== current.code && map[current.parent]) {
				map[current.parent].children.push(current);
			} else {
				roots.push(current);
			}
		}
		for (var k = 0; k < parsed.heads.length; k += 1) {
			var head = parsed.heads[k];
			if (map[head.parent]) {
				map[head.parent].heads.push(head);
			}
		}
		assignTreePaths(roots, "");
		return { roots: roots, map: map };
	}

	function assignTreePaths(items, parentPath) {
		for (var i = 0; i < items.length; i += 1) {
			var item = items[i];
			item.fullPath = parentPath ? parentPath + "/" + item.name : item.name;
			assignTreePaths(item.children || [], item.fullPath);
			for (var j = 0; j < (item.heads || []).length; j += 1) {
				item.heads[j].fullPath = item.fullPath ? item.fullPath + "/" + item.heads[j].name : item.heads[j].name;
			}
		}
	}

	function selectTreeNode(control, item, isHead) {
		var entry = treeControls[control.id];
		if (!entry) {
			return;
		}
		if (isHead) {
			var parent = entry.groups[item.parent] || {};
			entry.state.groupValue = item.parent || "0";
			entry.state.groupName = parent.name || "";
			entry.state.headValue = item.value || "0";
			entry.state.headName = item.name || "";
			entry.state.selectedKey = item.value || item.code || "";
			entry.state.selectedText = item.name || "";
			entry.state.selectedFullPath = item.fullPath || item.name || "";
			entry.state.classification = entry.state.selectedKey;
			entry.state.classificationName = entry.state.selectedText;
			control.value = entry.state.headValue;
		} else {
			entry.state.groupValue = item.code || "0";
			entry.state.groupName = item.name || "";
			entry.state.headValue = "0";
			entry.state.headName = "";
			entry.state.selectedKey = item.code || "";
			entry.state.selectedText = item.name || "";
			entry.state.selectedFullPath = item.fullPath || item.name || "";
			entry.state.classification = entry.state.selectedKey;
			entry.state.classificationName = entry.state.selectedText;
			control.value = entry.state.groupValue;
		}
		var selected = entry.host.querySelector(".itms-tree-selected");
		if (selected) {
			selected.className = selected.className.replace(/\bitms-tree-selected\b/g, "").replace(/\s+/g, " ").trim();
		}
		var activeClassName = String(document.activeElement && document.activeElement.className || "");
		if (document.activeElement && activeClassName.indexOf("itms-tree-label") !== -1) {
			document.activeElement.className += " itms-tree-selected";
		}
		control.dispatchEvent(new Event("change", { bubbles: true }));
	}

	function renderTreeNode(control, item, isHead) {
		var li = document.createElement("li");
		var row = document.createElement("div");
		var toggle = document.createElement("button");
		var label = document.createElement("button");
		var children = isHead ? [] : item.children.concat(item.heads);
		var childList;

		row.className = "itms-tree-row";
		toggle.type = "button";
		toggle.className = children.length ? "itms-tree-toggle" : "itms-tree-spacer";
		toggle.textContent = children.length ? "-" : "";
		toggle.setAttribute("aria-label", children.length ? "Toggle" : "");
		label.type = "button";
		label.className = isHead ? "itms-tree-label itms-tree-head" : "itms-tree-label itms-tree-group";
		label.textContent = item.name || item.code || item.value || "";
		label.onclick = function () {
			selectTreeNode(control, item, isHead);
		};
		row.appendChild(toggle);
		row.appendChild(label);
		li.appendChild(row);

		if (children.length) {
			childList = document.createElement("ul");
			for (var i = 0; i < item.children.length; i += 1) {
				childList.appendChild(renderTreeNode(control, item.children[i], false));
			}
			for (var j = 0; j < item.heads.length; j += 1) {
				childList.appendChild(renderTreeNode(control, item.heads[j], true));
			}
			toggle.onclick = function () {
				var collapsed = childList.style.display === "none";
				childList.style.display = collapsed ? "" : "none";
				toggle.textContent = collapsed ? "-" : "+";
			};
			li.appendChild(childList);
		}
		return li;
	}

	function renderTree(control, xmlDoc) {
		var entry = treeControls[control.id];
		var parsed = parseTreeXml(xmlDoc);
		var model = buildTreeModel(parsed);
		var title;
		var list;
		if (!entry) {
			return;
		}
		entry.groups = model.map;
		entry.host.innerHTML = "";
		if (entry.state.listName) {
			title = document.createElement("div");
			title.className = "itms-tree-title";
			title.textContent = entry.state.listName;
			entry.host.appendChild(title);
		}
		list = document.createElement("ul");
		for (var i = 0; i < model.roots.length; i += 1) {
			list.appendChild(renderTreeNode(control, model.roots[i], false));
		}
		if (!model.roots.length) {
			entry.host.innerHTML = '<div class="itms-tree-status">No records available</div>';
			return;
		}
		entry.host.appendChild(list);
	}

	function refreshTree(control) {
		var entry = treeControls[control.id];
		if (!entry || !entry.state.dsn) {
			if (entry) {
				entry.host.innerHTML = '<div class="itms-tree-status">No data source</div>';
			}
			return;
		}
		entry.host.innerHTML = '<div class="itms-tree-status">Loading...</div>';
		fetch(normalizeDataSource(resolveTreeDataSource(entry.state)), { cache: "no-cache", credentials: "same-origin" })
			.then(function (response) {
				if (!response.ok) {
					throw new Error(response.status + " " + response.statusText);
				}
				return response.text();
			})
			.then(function (text) {
				var xmlDoc = new DOMParser().parseFromString(text, "text/xml");
				if (xmlDoc.getElementsByTagName("parsererror").length) {
					throw new Error("Invalid XML");
				}
				renderTree(control, xmlDoc);
			})
			.catch(function (error) {
				entry.host.innerHTML = '<div class="itms-tree-status">Unable to load tree data</div>';
				if (window.console) {
					window.console.error("Tree load failed for " + control.id + ": " + error.message);
				}
			});
	}

	function resolveTreeDataSource(state) {
		var source = state.dsn || "";
		var iTypeParts;
		if (!/\.asp(?:[?#]|$)/i.test(source)) {
			source = source.replace(/[\\\/]*$/, "/");
			source += state.treeKind === "item-classification" ? "GetCategoryGroup.asp" : "GetGroup.asp";
		}
		if (state.treeKind === "item-classification") {
			iTypeParts = String(state.iType || "NO:NO:NO").split(":");
			source = withQueryParams(source, {
				sIT: iTypeParts[0] || "NO",
				sOrgID: iTypeParts[2] || "NO"
			});
		}
		return source;
	}

	function withQueryParams(source, params) {
		try {
			var url = new URL(source, window.location.href);
			Object.keys(params).forEach(function (name) {
				url.searchParams.set(name, params[name]);
			});
			return url.href;
		} catch (ignore) {
			var joiner = String(source).indexOf("?") === -1 ? "?" : "&";
			return String(source) + joiner + Object.keys(params).map(function (name) {
				return encodeURIComponent(name) + "=" + encodeURIComponent(params[name]);
			}).join("&");
		}
	}

	function normalizeDataSource(dsn) {
		try {
			var url = new URL(dsn, window.location.href);
			if (url.origin !== window.location.origin && url.pathname) {
				return url.pathname + url.search;
			}
			return url.href;
		} catch (ignore) {
			return dsn;
		}
	}

	function upgradeTrees(root) {
		var scope = root || document;
		var objects = scope.querySelectorAll("object[classid]");
		var placeholders = scope.querySelectorAll("[data-itms-tree-control]");
		for (var i = objects.length - 1; i >= 0; i -= 1) {
			if (isTreeObject(objects[i])) {
				createTreeControlFromObject(objects[i]);
			}
		}
		for (var j = placeholders.length - 1; j >= 0; j -= 1) {
			createTreeControlFromObject(placeholders[j]);
		}
	}

	function defineAlias(proto, name, descriptor) {
		try {
			if (proto && !(name in proto)) {
				Object.defineProperty(proto, name, descriptor);
			}
		} catch (ignore) {}
	}

	function defineDomPropertyAlias(proto, aliasName, realName) {
		defineAlias(proto, aliasName, {
			get: function () {
				if (realName in this) {
					return this[realName];
				}
				return this.getAttribute ? this.getAttribute(realName) : undefined;
			},
			set: function (value) {
				if (realName in this) {
					this[realName] = value;
				} else if (this.setAttribute) {
					this.setAttribute(realName, value == null ? "" : String(value));
				}
			}
		});
	}

	function defineDomMethodAlias(proto, aliasName, realName) {
		defineAlias(proto, aliasName, {
			value: function () {
				if (this[realName]) {
					return this[realName].apply(this, arguments);
				}
				return undefined;
			}
		});
	}

	function installDomPropertyAliases() {
		var elementProto = window.HTMLElement && HTMLElement.prototype || window.Element && Element.prototype;
		[
			["Value", "value"],
			["Checked", "checked"],
			["Selected", "selected"],
			["SelectedIndex", "selectedIndex"],
			["Disabled", "disabled"],
			["ReadOnly", "readOnly"],
			["ClassName", "className"],
			["InnerHTML", "innerHTML"],
			["InnerText", "innerText"],
			["Src", "src"],
			["Href", "href"],
			["Action", "action"],
			["Target", "target"],
			["Name", "name"],
			["Id", "id"],
			["Type", "type"],
			["Options", "options"],
			["Length", "length"]
		].forEach(function (entry) {
			defineDomPropertyAlias(elementProto, entry[0], entry[1]);
		});
		defineAlias(elementProto, "Style", {
			get: function () {
				return this.style;
			},
			set: function (value) {
				this.style.cssText = value == null ? "" : String(value);
			}
		});
		[
			["Focus", "focus"],
			["Select", "select"],
			["Submit", "submit"],
			["Reset", "reset"],
			["Click", "click"]
		].forEach(function (entry) {
			defineDomMethodAlias(elementProto, entry[0], entry[1]);
			if (window.HTMLFormElement) {
				defineDomMethodAlias(HTMLFormElement.prototype, entry[0], entry[1]);
			}
		});
		[
			window.NodeList && NodeList.prototype,
			window.HTMLCollection && HTMLCollection.prototype,
			window.HTMLFormControlsCollection && HTMLFormControlsCollection.prototype,
			window.HTMLOptionsCollection && HTMLOptionsCollection.prototype
		].forEach(function (proto) {
			defineDomPropertyAlias(proto, "Length", "length");
		});
	}

	function installEventCompatibilityAliases() {
		if (!window.__itmsEventShimInstalled) {
			window.__itmsEventShimInstalled = true;
			[
				"click",
				"dblclick",
				"mousedown",
				"mouseup",
				"mousemove",
				"mouseover",
				"mouseout",
				"keydown",
				"keyup",
				"keypress",
				"change",
				"input",
				"submit",
				"focus",
				"blur"
			].forEach(function (eventName) {
				window.addEventListener(eventName, function (evt) {
					window.__itmsCurrentEvent = evt;
				}, true);
			});
			defineAlias(window, "event", {
				get: function () {
					return window.__itmsCurrentEvent || null;
				},
				set: function (value) {
					window.__itmsCurrentEvent = value || null;
				}
			});
		}
		if (window.Event) {
			defineAlias(Event.prototype, "srcElement", {
				get: function () {
					return this.target || null;
				}
			});
			defineAlias(Event.prototype, "returnValue", {
				get: function () {
					return !this.defaultPrevented;
				},
				set: function (value) {
					if (value === false && this.preventDefault) {
						this.preventDefault();
					}
				}
			});
			defineAlias(Event.prototype, "cancelBubble", {
				get: function () {
					return false;
				},
				set: function (value) {
					if (value && this.stopPropagation) {
						this.stopPropagation();
					}
				}
			});
		}
		if (window.MouseEvent) {
			defineAlias(MouseEvent.prototype, "fromElement", {
				get: function () {
					return this.relatedTarget || null;
				}
			});
			defineAlias(MouseEvent.prototype, "toElement", {
				get: function () {
					return this.relatedTarget || null;
				}
			});
		}
		if (window.KeyboardEvent) {
			defineAlias(KeyboardEvent.prototype, "keyCode", {
				get: function () {
					if (typeof this.which === "number" && this.which !== 0) {
						return this.which;
					}
					if (this.key && this.key.length === 1) {
						return this.key.toUpperCase().charCodeAt(0);
					}
					return 0;
				}
			});
			defineAlias(KeyboardEvent.prototype, "which", {
				get: function () {
					return this.keyCode || 0;
				}
			});
		}
	}

	function serializeXml(node) {
		if (!node) {
			return "";
		}
		try {
			return new XMLSerializer().serializeToString(node);
		} catch (ignore) {
			return "";
		}
	}

	function parseXml(text) {
		var xmlText = String(text || "").trim();
		var doc;
		if (!xmlText) {
			xmlText = "<Root/>";
		}
		doc = new DOMParser().parseFromString(xmlText, "text/xml");
		if (doc.getElementsByTagName("parsererror").length) {
			doc = new DOMParser().parseFromString("<Root/>", "text/xml");
		}
		patchXmlDocument(doc);
		return doc;
	}

	function selectByXPath(context, expression, single) {
		var doc = context.nodeType === 9 ? context : context.ownerDocument;
		var resolver = createNamespaceResolver(doc);
		var result;
		var nodes = [];
		try {
			result = doc.evaluate(
				expression,
				context,
				resolver,
				single ? XPathResult.FIRST_ORDERED_NODE_TYPE : XPathResult.ORDERED_NODE_SNAPSHOT_TYPE,
				null
			);
		} catch (ignore) {
			return single ? null : nodes;
		}
		if (single) {
			return result.singleNodeValue;
		}
		for (var i = 0; i < result.snapshotLength; i += 1) {
			nodes.push(result.snapshotItem(i));
		}
		nodes.Item = function (index) {
			return this[index];
		};
		nodes.item = nodes.Item;
		return nodes;
	}

	function createNamespaceResolver(doc) {
		var namespaces = {};
		var props = doc && doc._itmsXmlProperties || {};
		var declarations = props.SelectionNamespaces || props.selectionnamespaces || "";
		String(declarations).replace(/xmlns(?::([A-Za-z_][\w.-]*))?\s*=\s*(['"])(.*?)\2/g, function (_, prefix, __, uri) {
			if (prefix) {
				namespaces[prefix] = uri;
			}
			return "";
		});
		return function (prefix) {
			if (Object.prototype.hasOwnProperty.call(namespaces, prefix)) {
				return namespaces[prefix];
			}
			if (doc && doc.documentElement && doc.documentElement.lookupNamespaceURI) {
				return doc.documentElement.lookupNamespaceURI(prefix) || null;
			}
			return null;
		};
	}

	function findAttributeCaseInsensitive(attributes, name) {
		var wanted;
		var attribute;
		if (!attributes || name == null) {
			return null;
		}
		wanted = String(name).toLowerCase();
		for (var i = 0; i < attributes.length; i += 1) {
			attribute = attributes.item(i);
			if (attribute && String(attribute.name || attribute.nodeName || "").toLowerCase() === wanted) {
				return attribute;
			}
		}
		return null;
	}

	function patchXmlDomAliases() {
		defineAlias(Node.prototype, "xml", {
			get: function () {
				return serializeXml(this);
			}
		});
		defineAlias(Document.prototype, "xml", {
			get: function () {
				return serializeXml(this);
			}
		});
		defineAlias(Node.prototype, "text", {
			get: function () {
				return this.textContent || "";
			},
			set: function (value) {
				this.textContent = value == null ? "" : String(value);
			}
		});
		defineAlias(Node.prototype, "Text", {
			get: function () {
				return this.textContent || "";
			},
			set: function (value) {
				this.textContent = value == null ? "" : String(value);
			}
		});
		defineAlias(Node.prototype, "nodeTypedValue", {
			get: function () {
				return this.nodeType === 2 ? this.value : this.textContent || "";
			},
			set: function (value) {
				if (this.nodeType === 2) {
					this.value = value == null ? "" : String(value);
				} else {
					this.textContent = value == null ? "" : String(value);
				}
			}
		});
		defineAlias(Element.prototype, "Attributes", {
			get: function () {
				return this.attributes;
			}
		});
		defineAlias(Attr.prototype, "Value", {
			get: function () {
				return this.value;
			},
			set: function (value) {
				this.value = value == null ? "" : String(value);
			}
		});
		defineAlias(Attr.prototype, "nodeTypedValue", {
			get: function () {
				return this.value;
			},
			set: function (value) {
				this.value = value == null ? "" : String(value);
			}
		});
		if (window.NamedNodeMap) {
			if (!NamedNodeMap.prototype.Item) {
				NamedNodeMap.prototype.Item = function (index) {
					return this.item(index);
				};
			}
			if (!NamedNodeMap.prototype._itmsCaseInsensitiveGetNamedItem) {
				try {
					(function () {
						var nativeGetNamedItem = NamedNodeMap.prototype.getNamedItem;
						NamedNodeMap.prototype.getNamedItem = function (name) {
							return nativeGetNamedItem.call(this, name) || findAttributeCaseInsensitive(this, name);
						};
						NamedNodeMap.prototype.GetNamedItem = NamedNodeMap.prototype.getNamedItem;
						NamedNodeMap.prototype._itmsCaseInsensitiveGetNamedItem = true;
					}());
				} catch (ignoreNamedNodeMapPatch) {}
			}
		}
		if (!Element.prototype._itmsCaseInsensitiveAttributes) {
			try {
				(function () {
					var nativeGetAttribute = Element.prototype.getAttribute;
					var nativeHasAttribute = Element.prototype.hasAttribute;
					var nativeSetAttribute = Element.prototype.setAttribute;
					Element.prototype.getAttribute = function (name) {
						var value = nativeGetAttribute.call(this, name);
						var attribute;
						if (value != null) {
							return value;
						}
						attribute = findAttributeCaseInsensitive(this.attributes, name);
						return attribute ? attribute.value : null;
					};
					Element.prototype.hasAttribute = function (name) {
						return nativeHasAttribute.call(this, name) || !!findAttributeCaseInsensitive(this.attributes, name);
					};
					Element.prototype.setAttribute = function (name, value) {
						var attribute = findAttributeCaseInsensitive(this.attributes, name);
						if (attribute) {
							attribute.value = value == null ? "" : String(value);
							return;
						}
						nativeSetAttribute.call(this, name, value);
					};
					Element.prototype._itmsCaseInsensitiveAttributes = true;
				}());
			} catch (ignoreElementAttributePatch) {}
		}
		if (window.NodeList && !NodeList.prototype.Item) {
			NodeList.prototype.Item = function (index) {
				return this.item(index);
			};
		}
		if (window.HTMLCollection && !HTMLCollection.prototype.Item) {
			HTMLCollection.prototype.Item = function (index) {
				return this.item(index);
			};
		}
		if (!Element.prototype.selectNodes) {
			Element.prototype.selectNodes = function (expression) {
				return selectByXPath(this, expression, false);
			};
		}
		if (!Element.prototype.selectSingleNode) {
			Element.prototype.selectSingleNode = function (expression) {
				return selectByXPath(this, expression, true);
			};
		}
		if (!Document.prototype.selectNodes) {
			Document.prototype.selectNodes = function (expression) {
				return selectByXPath(this, expression, false);
			};
		}
		if (!Document.prototype.selectSingleNode) {
			Document.prototype.selectSingleNode = function (expression) {
				return selectByXPath(this, expression, true);
			};
		}
		if (!Document.prototype.setProperty) {
			Document.prototype.setProperty = function (name, value) {
				return setXmlDocumentProperty(this, name, value);
			};
		}
		if (!Document.prototype.getProperty) {
			Document.prototype.getProperty = function (name) {
				return getXmlDocumentProperty(this, name);
			};
		}
		if (!Document.prototype.transformNode) {
			Document.prototype.transformNode = function (stylesheet) {
				return transformXmlNode(this, stylesheet);
			};
		}
		if (!Document.prototype.transformNodeToObject) {
			Document.prototype.transformNodeToObject = function (stylesheet, outputDoc) {
				return transformXmlNodeToObject(this, stylesheet, outputDoc);
			};
		}
	}

	function setXmlDocumentProperty(doc, name, value) {
		var key;
		if (!doc || name == null) {
			return value;
		}
		key = String(name);
		doc._itmsXmlProperties = doc._itmsXmlProperties || {};
		doc._itmsXmlProperties[key] = value;
		doc._itmsXmlProperties[key.toLowerCase()] = value;
		return value;
	}

	function getXmlDocumentProperty(doc, name) {
		var key;
		var props;
		if (!doc || name == null) {
			return undefined;
		}
		key = String(name);
		props = doc._itmsXmlProperties || {};
		if (Object.prototype.hasOwnProperty.call(props, key)) {
			return props[key];
		}
		return props[key.toLowerCase()];
	}

	function defineXmlDocumentValue(doc, name, initialValue) {
		try {
			if (Object.prototype.hasOwnProperty.call(doc, name)) {
				return;
			}
			Object.defineProperty(doc, name, {
				configurable: true,
				get: function () {
					var value = getXmlDocumentProperty(doc, name);
					return value === undefined ? initialValue : value;
				},
				set: function (value) {
					setXmlDocumentProperty(doc, name, value);
				}
			});
		} catch (ignore) {}
	}

	function normalizeXmlDocument(value) {
		if (!value) {
			return null;
		}
		if (value.XMLDocument) {
			return value.XMLDocument;
		}
		if (value._doc) {
			return value._doc;
		}
		if (value.nodeType === 9) {
			return value;
		}
		if (value.nodeType && value.ownerDocument) {
			return value.ownerDocument;
		}
		if (typeof value === "string") {
			return parseXml(value);
		}
		return null;
	}

	function transformXmlNode(source, stylesheet) {
		var styleDoc = normalizeXmlDocument(stylesheet);
		var processor;
		var result;
		if (!window.XSLTProcessor || !source || !styleDoc || !styleDoc.documentElement) {
			return "";
		}
		try {
			processor = new XSLTProcessor();
			processor.importStylesheet(styleDoc);
			result = processor.transformToDocument(source);
			return serializeXml(result);
		} catch (ignore) {
			return "";
		}
	}

	function transformXmlNodeToObject(source, stylesheet, outputDoc) {
		var xmlText = transformXmlNode(source, stylesheet);
		var target = normalizeXmlDocument(outputDoc);
		if (!target || !xmlText) {
			return false;
		}
		if (target.loadXML) {
			return target.loadXML(xmlText);
		}
		return false;
	}

	function createXmlNode(doc, type, name, namespaceUri) {
		var nodeType = type;
		var nodeName = name;
		if (name == null) {
			nodeName = type;
			nodeType = 1;
		}
		nodeType = typeof nodeType === "number" ? nodeType : String(nodeType || "element").toLowerCase();
		if (nodeType === 1 || nodeType === "1" || nodeType === "element" || nodeType === "node_element") {
			return namespaceUri ? doc.createElementNS(namespaceUri, nodeName) : doc.createElement(nodeName);
		}
		if (nodeType === 2 || nodeType === "2" || nodeType === "attribute" || nodeType === "node_attribute") {
			return namespaceUri ? doc.createAttributeNS(namespaceUri, nodeName) : doc.createAttribute(nodeName);
		}
		if (nodeType === 3 || nodeType === "3" || nodeType === "text" || nodeType === "node_text") {
			return doc.createTextNode(nodeName || "");
		}
		if (nodeType === 4 || nodeType === "4" || nodeType === "cdata" || nodeType === "node_cdata_section") {
			return doc.createCDATASection(nodeName || "");
		}
		if (nodeType === 8 || nodeType === "8" || nodeType === "comment" || nodeType === "node_comment") {
			return doc.createComment(nodeName || "");
		}
		return namespaceUri ? doc.createElementNS(namespaceUri, nodeName) : doc.createElement(nodeName);
	}

	function patchXmlDocument(doc) {
		if (!doc || doc._itmsXmlPatched) {
			return doc;
		}
		try {
			doc._itmsXmlPatched = true;
			doc._itmsXmlProperties = doc._itmsXmlProperties || {};
			defineXmlDocumentValue(doc, "async", false);
			defineXmlDocumentValue(doc, "preserveWhiteSpace", false);
			defineXmlDocumentValue(doc, "validateOnParse", false);
			defineXmlDocumentValue(doc, "resolveExternals", false);
			defineXmlDocumentValue(doc, "readyState", 4);
			defineXmlDocumentValue(doc, "parseError", {
				errorCode: 0,
				reason: "",
				line: 0,
				linepos: 0,
				srcText: ""
			});
			doc.setProperty = function (name, value) {
				return setXmlDocumentProperty(this, name, value);
			};
			doc.getProperty = function (name) {
				return getXmlDocumentProperty(this, name);
			};
			doc.loadXML = function (text) {
				var newDoc = parseXml(text);
				while (this.firstChild) {
					this.removeChild(this.firstChild);
				}
				if (newDoc.documentElement) {
					this.appendChild(this.importNode(newDoc.documentElement, true));
				}
				return true;
			};
			doc.load = function (url) {
				var xhr = new XMLHttpRequest();
				xhr.open("GET", normalizeDataSource(url), false);
				xhr.send(null);
				if (xhr.status === 0 || (xhr.status >= 200 && xhr.status < 300)) {
					return this.loadXML(xhr.responseText);
				}
				return false;
			};
			doc.LoadXML = doc.loadXML;
			doc.Load = doc.load;
			doc.CreateElement = function (name) {
				return this.createElement(name);
			};
			doc.CreateAttribute = function (name) {
				return this.createAttribute(name);
			};
			doc.CreateTextNode = function (text) {
				return this.createTextNode(text || "");
			};
			doc.CreateCDATASection = function (text) {
				return this.createCDATASection(text || "");
			};
			doc.createNode = function (type, name, namespaceUri) {
				return createXmlNode(this, type, name, namespaceUri);
			};
			doc.CreateNode = doc.createNode;
			doc.transformNode = function (stylesheet) {
				return transformXmlNode(this, stylesheet);
			};
			doc.transformNodeToObject = function (stylesheet, outputDoc) {
				return transformXmlNodeToObject(this, stylesheet, outputDoc);
			};
		} catch (ignore) {}
		return doc;
	}

	function exposeXmlIslandOnElement(element, wrapper) {
		var descriptors = {
			XMLDocument: {
				get: function () {
					return wrapper.XMLDocument;
				}
			},
			documentElement: {
				get: function () {
					return wrapper.documentElement;
				}
			},
			xml: {
				get: function () {
					return wrapper.xml;
				}
			},
			XML: {
				get: function () {
					return wrapper.XML;
				}
			},
			childNodes: {
				get: function () {
					return wrapper.childNodes;
				}
			},
			firstChild: {
				get: function () {
					return wrapper.firstChild;
				}
			},
			lastChild: {
				get: function () {
					return wrapper.lastChild;
				}
			},
			nodeType: {
				get: function () {
					return wrapper.nodeType;
				}
			},
			nodeName: {
				get: function () {
					return wrapper.nodeName;
				}
			},
			nodeValue: {
				get: function () {
					return wrapper.nodeValue;
				}
			},
			nodeTypedValue: {
				get: function () {
					return wrapper.nodeTypedValue;
				}
			},
			readyState: {
				get: function () {
					return wrapper.readyState;
				}
			},
			parseError: {
				get: function () {
					return wrapper.parseError;
				}
			}
		};
		Object.keys(descriptors).forEach(function (name) {
			try {
				Object.defineProperty(element, name, {
					configurable: true,
					get: descriptors[name].get
				});
			} catch (ignorePropertyBridge) {}
		});
		[
			"createElement",
			"CreateElement",
			"createAttribute",
			"CreateAttribute",
			"createTextNode",
			"CreateTextNode",
			"createCDATASection",
			"CreateCDATASection",
			"createNode",
			"CreateNode",
			"appendChild",
			"AppendChild",
			"removeChild",
			"RemoveChild",
			"replaceChild",
			"ReplaceChild",
			"insertBefore",
			"InsertBefore",
			"cloneNode",
			"CloneNode",
			"importNode",
			"ImportNode",
			"hasChildNodes",
			"getElementsByTagName",
			"GetElementsByTagName",
			"loadXML",
			"LoadXML",
			"load",
			"Load",
			"selectNodes",
			"selectSingleNode",
			"setProperty",
			"getProperty",
			"transformNode",
			"transformNodeToObject"
		].forEach(function (name) {
			try {
				element[name] = function () {
					return wrapper[name].apply(wrapper, arguments);
				};
			} catch (ignoreMethodBridge) {}
		});
	}

	function createXmlIsland(element) {
		var source = element.getAttribute("data-src") || element.getAttribute("src");
		var text = String(element.tagName || "").toLowerCase() === "script" ? element.text || element.textContent || "" : element.innerHTML || element.textContent || "";
		var doc = parseXml(text);
		var wrapper = {
			_element: element,
			_doc: doc,
			get XMLDocument() {
				return this._doc;
			},
			get documentElement() {
				return this._doc.documentElement;
			},
			get xml() {
				return serializeXml(this._doc);
			},
			get XML() {
				return serializeXml(this._doc);
			},
			get childNodes() {
				return this._doc.childNodes;
			},
			get firstChild() {
				return this._doc.firstChild;
			},
			get lastChild() {
				return this._doc.lastChild;
			},
			get nodeType() {
				return this._doc.nodeType;
			},
			get nodeName() {
				return this._doc.nodeName;
			},
			get nodeValue() {
				return this._doc.nodeValue;
			},
			get nodeTypedValue() {
				return this._doc.nodeTypedValue;
			},
			get readyState() {
				return this._doc.readyState;
			},
			get parseError() {
				return this._doc.parseError;
			},
			createElement: function (name) {
				return this._doc.createElement(name);
			},
			CreateElement: function (name) {
				return this.createElement(name);
			},
			createAttribute: function (name) {
				return this._doc.createAttribute(name);
			},
			CreateAttribute: function (name) {
				return this.createAttribute(name);
			},
			createTextNode: function (text) {
				return this._doc.createTextNode(text || "");
			},
			CreateTextNode: function (text) {
				return this.createTextNode(text);
			},
			createCDATASection: function (text) {
				return this._doc.createCDATASection(text || "");
			},
			CreateCDATASection: function (text) {
				return this.createCDATASection(text);
			},
			createNode: function (type, name, namespaceUri) {
				return this._doc.createNode(type, name, namespaceUri);
			},
			CreateNode: function (type, name, namespaceUri) {
				return this.createNode(type, name, namespaceUri);
			},
			appendChild: function (node) {
				return this._doc.appendChild(node);
			},
			AppendChild: function (node) {
				return this.appendChild(node);
			},
			removeChild: function (node) {
				return this._doc.removeChild(node);
			},
			RemoveChild: function (node) {
				return this.removeChild(node);
			},
			replaceChild: function (newChild, oldChild) {
				return this._doc.replaceChild(newChild, oldChild);
			},
			ReplaceChild: function (newChild, oldChild) {
				return this.replaceChild(newChild, oldChild);
			},
			insertBefore: function (newChild, refChild) {
				return this._doc.insertBefore(newChild, refChild || null);
			},
			InsertBefore: function (newChild, refChild) {
				return this.insertBefore(newChild, refChild);
			},
			cloneNode: function (deep) {
				return this._doc.cloneNode(deep !== false);
			},
			CloneNode: function (deep) {
				return this.cloneNode(deep);
			},
			importNode: function (node, deep) {
				return this._doc.importNode(node, deep !== false);
			},
			ImportNode: function (node, deep) {
				return this.importNode(node, deep);
			},
			hasChildNodes: function () {
				return this._doc.hasChildNodes();
			},
			getElementsByTagName: function (name) {
				return this._doc.getElementsByTagName(name);
			},
			GetElementsByTagName: function (name) {
				return this.getElementsByTagName(name);
			},
			loadXML: function (text) {
				this._doc = parseXml(text);
				return true;
			},
			LoadXML: function (text) {
				return this.loadXML(text);
			},
			load: function (url) {
				var xhr = new XMLHttpRequest();
				xhr.open("GET", normalizeDataSource(url), false);
				xhr.send(null);
				if (xhr.status === 0 || (xhr.status >= 200 && xhr.status < 300)) {
					return this.loadXML(xhr.responseText);
				}
				return false;
			},
			Load: function (url) {
				return this.load(url);
			},
			selectNodes: function (expression) {
				return this._doc.selectNodes(expression);
			},
			selectSingleNode: function (expression) {
				return this._doc.selectSingleNode(expression);
			},
			setProperty: function (name, value) {
				return this._doc.setProperty(name, value);
			},
			getProperty: function (name) {
				return this._doc.getProperty(name);
			},
			transformNode: function (stylesheet) {
				return this._doc.transformNode(stylesheet);
			},
			transformNodeToObject: function (stylesheet, outputDoc) {
				return this._doc.transformNodeToObject(stylesheet, outputDoc);
			}
		};
		if (source) {
			wrapper.load(source);
		}
		element.style.display = "none";
		element._itmsXmlIsland = wrapper;
		exposeXmlIslandOnElement(element, wrapper);
		return wrapper;
	}

	function upgradeXmlIslands(root) {
		var scope = root || document;
		var islands = scope.querySelectorAll("xml[id], XML[id], script[data-itms-xml-island][id]");
		var island;
		var id;
		patchXmlDomAliases();
		for (var i = 0; i < islands.length; i += 1) {
			id = islands[i].getAttribute("id");
			if (!id || islands[i]._itmsXmlIsland) {
				continue;
			}
			island = createXmlIsland(islands[i]);
			window[id] = island;
			document[id] = island;
		}
	}

	function namedDocumentItem(name) {
		if (!name) {
			return null;
		}
		return document.forms[name] || document.getElementsByName(name)[0] || document.getElementById(name) || null;
	}

	function aliasNameVariants(name) {
		var text = String(name || "");
		var variants = {};
		function add(value) {
			if (value) {
				variants[value] = true;
			}
		}
		add(text);
		add(text.toLowerCase());
		add(text.toUpperCase());
		add(text.charAt(0).toLowerCase() + text.slice(1));
		add(text.charAt(0).toUpperCase() + text.slice(1));
		add(text.replace(/Id\b/g, "ID"));
		add(text.replace(/ID\b/g, "Id"));
		return Object.keys(variants);
	}

	function defineObjectGetterAlias(object, name, getter) {
		try {
			if (object && name && !(name in object)) {
				Object.defineProperty(object, name, {
					configurable: true,
					get: getter
				});
			}
		} catch (ignore) {}
	}

	function defineDocumentNamedAlias(name) {
		defineObjectGetterAlias(document, name, function () {
			return namedDocumentItem(name);
		});
	}

	function namedCollectionItem(collection, name) {
		var item = null;
		if (!collection || !name) {
			return null;
		}
		try {
			item = collection.namedItem ? collection.namedItem(name) : collection[name];
		} catch (ignoreNamedItem) {}
		if (item) {
			return item;
		}
		name = String(name).toLowerCase();
		for (var i = 0; i < collection.length; i += 1) {
			item = collection[i];
			if (item && (String(item.name || "").toLowerCase() === name || String(item.id || "").toLowerCase() === name)) {
				return item;
			}
		}
		return null;
	}

	function namedFormControl(form, name) {
		var elements = form && form.elements;
		var matches = [];
		var item;
		if (!elements || !name) {
			return null;
		}
		try {
			item = elements.namedItem ? elements.namedItem(name) : elements[name];
		} catch (ignoreNamedControl) {}
		if (item) {
			return item;
		}
		name = String(name).toLowerCase();
		for (var i = 0; i < elements.length; i += 1) {
			item = elements[i];
			if (item && (String(item.name || "").toLowerCase() === name || String(item.id || "").toLowerCase() === name)) {
				matches.push(item);
			}
		}
		if (matches.length > 1) {
			matches.item = function (index) {
				return this[index];
			};
			return matches;
		}
		return matches[0] || null;
	}

	function defineFormAlias(form, aliasName, lookupName) {
		defineObjectGetterAlias(form, aliasName, function () {
			return namedFormControl(form, lookupName);
		});
	}

	function defineFormsCollectionAlias(aliasName, lookupName) {
		defineObjectGetterAlias(document.forms, aliasName, function () {
			return namedCollectionItem(document.forms, lookupName);
		});
	}

	function installLegacyFormControlAliases() {
		var forms = document.forms || [];
		var aliases;
		var element;
		for (var i = 0; i < forms.length; i += 1) {
			if (forms[i].name || forms[i].id) {
				aliasNameVariants(forms[i].name || forms[i].id).forEach(function (alias) {
					defineFormsCollectionAlias(alias, forms[i].name || forms[i].id);
				});
			}
			for (var j = 0; j < forms[i].elements.length; j += 1) {
				element = forms[i].elements[j];
				aliases = aliasNameVariants(element.name || element.id);
				aliases.forEach(function (alias) {
					defineFormAlias(forms[i], alias, element.name || element.id);
				});
			}
		}
	}

	function installLegacyNamedAliases() {
		var forms = document.forms || [];
		["formname", "FormName", "frm1", "form1", "Form1", "form1name", "form1id"].forEach(function (name) {
			defineDocumentNamedAlias(name);
		});
		for (var i = 0; i < forms.length; i += 1) {
			if (forms[i].name) {
				defineDocumentNamedAlias(forms[i].name);
			}
			if (forms[i].id) {
				defineDocumentNamedAlias(forms[i].id);
			}
		}
		installLegacyFormControlAliases();
	}

	function installCreateObjectShim() {
		if (window.CreateObject) {
			return;
		}
		window.CreateObject = function (progId) {
			var name = String(progId || "").toLowerCase();
			if (name.indexOf("xmlhttp") !== -1) {
				return new XMLHttpRequest();
			}
			if (name.indexOf("xmldom") !== -1) {
				return patchXmlDocument(parseXml("<Root/>"));
			}
			throw new Error("Unsupported ActiveX CreateObject: " + progId);
		};
	}

	function installUtilityShims() {
		if (!window.Cancel) {
			window.Cancel = function (sLoc) {
				if (confirm("Do you want to Cancel, If so the data entered will be lost.")) {
					window.location.href = sLoc;
				}
			};
		}
		if (!window.GetWindowSizeForPopup) {
			window.GetWindowSizeForPopup = function (sPopupType) {
				var sizes = {
					1: "ItemSelectRelPartyCommon.asp:500:850",
					2: "PartySelection.asp:500:500",
					3: "DynamicNoSelection.asp:500:850",
					4: "MisPartySelection.asp:500:350",
					5: "GLHeadSelection.asp:500:350",
					6: "ItemSelectCommonForQuote.asp:500:850",
					7: "PartySelectionWithParTypeSel.asp:500:500",
					8: "EmpSelPop.asp:500:500",
					9: "PartySelPop.asp:500:500",
					10: "SupplierItemSelectCommon.asp:500:850",
					11: "PackingLotSerialDetails.asp:460:650",
					12: "PartySelectionAcc.asp:500:500"
				};
				return sizes[Number(sPopupType)] || "";
			};
		}
		if (!window.PrintWindow) {
			window.PrintWindow = function (sPara) {
				window.open(sPara, "PrintWindow", "height=200,width=300,resizable=no,status=no");
			};
		}
		if (!window.Paginate) {
			window.Paginate = function (iPageNo) {
				var form = document.forms.formname || document.forms["formname"] || document.formname || document.forms[0] || null;
				if (!form) {
					return;
				}
				form.hPageSelection.value = iPageNo;
				form.submit();
			};
		}
		if (!window.PaginateAcc) {
			window.PaginateAcc = function (iPageNo) {
				var form = document.forms.formname || document.forms["formname"] || document.formname || document.forms[0] || null;
				if (!form) {
					return;
				}
				form.hPageSelection.value = iPageNo;
				if (typeof window.GetFormDet === "function") {
					window.GetFormDet();
				}
				form.submit();
			};
		}
		if (!window.RndOff) {
			window.RndOff = function (iValue) {
				var value = Number(iValue);
				return isNaN(value) ? 0 : Math.round(value);
			};
		}
		if (!window.getDate) {
			window.getDate = function () {
				return new Date().toLocaleDateString();
			};
		}
		if (!window.getDec4) {
			window.getDec4 = function (value) {
				return Math.round(Number(value) * 10000) / 10000;
			};
		}
		if (!window.getDec6) {
			window.getDec6 = function (value) {
				return Math.round(Number(value) * 1000000) / 1000000;
			};
		}
	}

	function parseDialogFeatures(features) {
		var result = {};
		var text = String(features || "");
		var parts = text.split(";");
		var match;
		for (var i = 0; i < parts.length; i += 1) {
			match = parts[i].match(/^\s*([^:]+)\s*:\s*(.*?)\s*$/);
			if (match) {
				result[match[1].toLowerCase()] = match[2];
			}
		}
		return result;
	}

	function dialogWindowFeatures(features) {
		var parsed = parseDialogFeatures(features);
		var windowFeatures = [];
		if (parsed.dialogwidth) {
			windowFeatures.push("width=" + parseInt(parsed.dialogwidth, 10));
		}
		if (parsed.dialogheight) {
			windowFeatures.push("height=" + parseInt(parsed.dialogheight, 10));
		}
		if (String(parsed.resizable || "").toLowerCase() === "yes") {
			windowFeatures.push("resizable=yes");
		} else {
			windowFeatures.push("resizable=no");
		}
		if (String(parsed.status || "").toLowerCase() === "yes") {
			windowFeatures.push("status=yes");
		} else {
			windowFeatures.push("status=no");
		}
		windowFeatures.push("scrollbars=yes");
		return windowFeatures.join(",");
	}

	function addDialogId(url, id) {
		var joiner = String(url).indexOf("?") === -1 ? "?" : "&";
		return String(url) + joiner + "__itmsDialogId=" + encodeURIComponent(id);
	}

	function getDialogId() {
		var match = window.location.search.match(/[?&]__itmsDialogId=([^&]+)/);
		return match ? decodeURIComponent(match[1]) : "";
	}

	function receiveDialogValue(id, value) {
		var entry = window.__itmsDialogCallbacks && window.__itmsDialogCallbacks[id];
		if (!entry) {
			return;
		}
		delete window.__itmsDialogCallbacks[id];
		if (window.__itmsDialogArgs) {
			delete window.__itmsDialogArgs[id];
		}
		if (entry.timer) {
			window.clearInterval(entry.timer);
		}
		entry.callback(value);
	}

	function receiveDialogMessage(event) {
		var data = event && event.data;
		var entry;
		if (!data || data.type !== "itms-dialog-return" || !data.id) {
			return;
		}
		entry = window.__itmsDialogCallbacks && window.__itmsDialogCallbacks[data.id];
		if (entry && entry.popup && event.source && event.source !== entry.popup) {
			return;
		}
		receiveDialogValue(data.id, data.value);
	}

	function notifyDialogValue(id, value) {
		if (!id || !window.opener) {
			return;
		}
		try {
			if (window.opener.ITMSModernCompat && window.opener.ITMSModernCompat._receiveDialogValue) {
				window.opener.ITMSModernCompat._receiveDialogValue(id, value);
				return;
			}
		} catch (ignoreDirectReturn) {}
		try {
			window.opener.postMessage({
				type: "itms-dialog-return",
				id: id,
				value: value
			}, window.location.origin || "*");
		} catch (ignoreMessageReturn) {}
	}

	function returnModalValue(value) {
		var id = getDialogId();
		var returnedValue;
		if (value !== undefined) {
			window.returnValue = value;
			window.returnvalue = value;
		}
		returnedValue = window.returnValue !== undefined ? window.returnValue : window.returnvalue;
		if (returnedValue === undefined) {
			return;
		}
		notifyDialogValue(id, returnedValue);
	}

	function installDialogBridge() {
		var id = getDialogId();
		var nativeClose;
		if (!window.__itmsDialogCallbacks) {
			window.__itmsDialogCallbacks = {};
		}
		if (!window.__itmsDialogArgs) {
			window.__itmsDialogArgs = {};
		}
		if (!window.__itmsDialogMessagePatched) {
			window.__itmsDialogMessagePatched = true;
			window.addEventListener("message", receiveDialogMessage);
		}
		if (!window.ITMSModernCompatOpenDialog) {
			window.ITMSModernCompatOpenDialog = function (url, args, features, callback) {
				var dialogId = "dlg" + String(new Date().getTime()) + String(Math.floor(Math.random() * 1000000));
				var popup = window.open(addDialogId(url, dialogId), "_blank", dialogWindowFeatures(features));
				if (!popup) {
					alert("Popup was blocked. Please allow popups for this site and try again.");
					return null;
				}
				window.__itmsDialogArgs[dialogId] = args;
				window.__itmsDialogCallbacks[dialogId] = {
					callback: typeof callback === "function" ? callback : function () {},
					popup: popup,
					timer: window.setInterval(function () {
						var returnedValue;
						if (popup.closed) {
							try {
								returnedValue = popup.returnValue !== undefined ? popup.returnValue : popup.returnvalue;
							} catch (ignoreClosedPopup) {}
							receiveDialogValue(dialogId, returnedValue);
						}
					}, 500)
				};
				try {
					popup.dialogArguments = args;
				} catch (ignore) {}
				return popup;
			};
	}

		if (id && window.opener && window.opener.__itmsDialogArgs && id in window.opener.__itmsDialogArgs) {
			window.dialogArguments = window.opener.__itmsDialogArgs[id];
		}
		if (id && !window.__itmsDialogClosePatched) {
			window.__itmsDialogClosePatched = true;
			nativeClose = window.close;
			window.close = function () {
				returnModalValue();
				nativeClose.call(window);
			};
			window.addEventListener("beforeunload", function () {
				returnModalValue();
			});
		}
	}

	function syncDatePickersBeforeSubmit() {
		var forms = document.forms;
		for (var i = 0; i < forms.length; i += 1) {
			if (forms[i].getAttribute("data-itms-date-submit-ready") === "1") {
				continue;
			}
			forms[i].setAttribute("data-itms-date-submit-ready", "1");
			forms[i].addEventListener("submit", function () {
				var inputs = this.querySelectorAll("input.itms-date-picker");
				for (var j = 0; j < inputs.length; j += 1) {
					syncRelatedHiddenFields(inputs[j]);
				}
			});
		}
	}

	function ensureStyles() {
		if (document.getElementById("itms-modern-compat-style")) {
			return;
		}
		var style = document.createElement("style");
		style.id = "itms-modern-compat-style";
		style.textContent =
			"input.itms-date-picker{box-sizing:border-box;min-width:8.5em;height:20px;font:inherit;}" +
			".itms-tree{box-sizing:border-box;overflow:auto;border:1px solid #9a9a9a;background:#fff;color:#111;font:12px Arial,sans-serif;}" +
			".itms-tree-title{padding:4px 6px;font-weight:bold;background:#ececec;border-bottom:1px solid #c8c8c8;}" +
			".itms-tree ul{list-style:none;margin:0;padding-left:14px;}" +
			".itms-tree>ul{padding:4px;}" +
			".itms-tree-row{display:flex;align-items:center;min-height:20px;white-space:nowrap;}" +
			".itms-tree-toggle,.itms-tree-spacer{width:18px;height:18px;margin:0 2px 0 0;padding:0;border:0;background:transparent;line-height:18px;}" +
			".itms-tree-toggle{cursor:pointer;font-weight:bold;}" +
			".itms-tree-label{border:0;background:transparent;padding:2px 4px;text-align:left;cursor:pointer;font:inherit;}" +
			".itms-tree-label:hover,.itms-tree-selected{background:#dbeafe;}" +
			".itms-tree-head{color:#003c78;}" +
			".itms-tree-status{padding:8px;color:#555;}";
		(document.head || document.getElementsByTagName("head")[0] || document.documentElement).appendChild(style);
	}

	function init(root) {
		ensureStyles();
		installDomPropertyAliases();
		installEventCompatibilityAliases();
		installLegacyNamedAliases();
		installCreateObjectShim();
		installUtilityShims();
		installDialogBridge();
		upgradeXmlIslands(root || document);
		upgradeDatePickers(root || document);
		upgradeTrees(root || document);
		syncDatePickersBeforeSubmit();
	}

	window.ITMSModernCompat = {
		init: init,
		upgradeDatePickers: upgradeDatePickers,
		upgradeTrees: upgradeTrees,
		upgradeXmlIslands: upgradeXmlIslands,
		decorateDateInput: decorateDateInput,
		toDisplayDate: toDisplayDate,
		toIsoDate: toIsoDate,
		openModalDialog: function (url, args, features, callback) {
			return window.ITMSModernCompatOpenDialog(url, args, features, callback);
		},
		returnModalValue: returnModalValue,
		_receiveDialogValue: receiveDialogValue,
		getTree: function (id) {
			return treeControls[id] && treeControls[id].control;
		}
	};

	init(document);
	onReady(function () {
		init(document);
	});
}());
