(function () {
	"use strict";

	var DATE_PICKER_CLSID = "01e5bf20-f919-44e6-a698-cf7fd7c7d6cd";
	var TREE_CLSID = "355ceafa-cb06-4345-8384-d0725c8a3048";
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

	function readParams(objectEl) {
		var params = {};
		var nodes = objectEl.getElementsByTagName("param");
		for (var i = 0; i < nodes.length; i += 1) {
			var name = nodes[i].getAttribute("name");
			if (name) {
				params[name.toLowerCase()] = nodes[i].getAttribute("value") || "";
			}
		}
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
			headName: params.headname || ""
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
		host.style.width = objectEl.getAttribute("width") || "263px";
		host.style.height = objectEl.getAttribute("height") || "353px";
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
		control.Refresh = function () {
			refreshTree(control);
		};
		control.refresh = control.Refresh;
		refreshTree(control);
		return control;
	}

	function parseTreeXml(xmlDoc) {
		var groups = [];
		var heads = [];
		var elements = xmlDoc.getElementsByTagName("*");
		for (var i = 0; i < elements.length; i += 1) {
			var tagName = elements[i].tagName || "";
			if (/group$/i.test(tagName) && elements[i].getAttribute("GroupCode") != null) {
				groups.push({
					code: elements[i].getAttribute("GroupCode") || "0",
					name: elements[i].getAttribute("GroupName") || "",
					parent: elements[i].getAttribute("ParentCode") || ""
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
		return { roots: roots, map: map };
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
			control.value = entry.state.headValue;
		} else {
			entry.state.groupValue = item.code || "0";
			entry.state.groupName = item.name || "";
			entry.state.headValue = "0";
			entry.state.headName = "";
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
		fetch(normalizeDataSource(entry.state.dsn), { cache: "no-cache", credentials: "same-origin" })
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
		for (var i = objects.length - 1; i >= 0; i -= 1) {
			if (hasClassId(objects[i], TREE_CLSID)) {
				createTreeControlFromObject(objects[i]);
			}
		}
	}

	function defineAlias(proto, name, descriptor) {
		try {
			if (proto && !(name in proto)) {
				Object.defineProperty(proto, name, descriptor);
			}
		} catch (ignore) {}
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
		var result;
		var nodes = [];
		try {
			result = doc.evaluate(
				expression,
				context,
				null,
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
		if (window.NamedNodeMap && !NamedNodeMap.prototype.Item) {
			NamedNodeMap.prototype.Item = function (index) {
				return this.item(index);
			};
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
	}

	function patchXmlDocument(doc) {
		if (!doc || doc._itmsXmlPatched) {
			return doc;
		}
		try {
			doc._itmsXmlPatched = true;
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
		} catch (ignore) {}
		return doc;
	}

	function createXmlIsland(element) {
		var source = element.getAttribute("src");
		var doc = parseXml(element.innerHTML || element.textContent || "");
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
			appendChild: function (node) {
				return this._doc.appendChild(node);
			},
			AppendChild: function (node) {
				return this.appendChild(node);
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
			}
		};
		if (source) {
			wrapper.load(source);
		}
		element.style.display = "none";
		element._itmsXmlIsland = wrapper;
		try {
			element.XMLDocument = wrapper.XMLDocument;
		} catch (ignore) {}
		return wrapper;
	}

	function upgradeXmlIslands(root) {
		var scope = root || document;
		var islands = scope.querySelectorAll("xml[id], XML[id]");
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
				document.formname.hPageSelection.value = iPageNo;
				document.formname.submit();
			};
		}
		if (!window.PaginateAcc) {
			window.PaginateAcc = function (iPageNo) {
				document.formname.hPageSelection.value = iPageNo;
				if (typeof window.GetFormDet === "function") {
					window.GetFormDet();
				}
				document.formname.submit();
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
		if (entry.timer) {
			window.clearInterval(entry.timer);
		}
		entry.callback(value);
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
		if (id && window.opener && window.opener.ITMSModernCompat && window.opener.ITMSModernCompat._receiveDialogValue) {
			window.opener.ITMSModernCompat._receiveDialogValue(id, returnedValue);
		}
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
		document.head.appendChild(style);
	}

	function init(root) {
		ensureStyles();
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

	onReady(function () {
		init(document);
	});
}());
