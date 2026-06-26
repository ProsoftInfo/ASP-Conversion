(function (window, document) {
	"use strict";

	var config = window.__itmsPopupCompat || {};
	var sent = false;
	var defaultSelectionReturn = "0/0";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form(name) {
		return document.forms[name || config.formName || "formname"] || document.forms[0] || null;
	}

	function field(name, formName) {
		var frm = form(formName);
		return frm && frm.elements ? frm.elements[name] || null : null;
	}

	function fields(name, formName) {
		var item = field(name, formName);
		if (!item) {
			return [];
		}
		if (item.length != null && !item.tagName) {
			return Array.prototype.slice.call(item);
		}
		if ((window.RadioNodeList && item instanceof RadioNodeList) || (window.NodeList && item instanceof NodeList) || (window.HTMLCollection && item instanceof HTMLCollection)) {
			return Array.prototype.slice.call(item);
		}
		return [item];
	}

	function byId(id) {
		return document.getElementById(id) || document.getElementsByName(id)[0] || window[id] || null;
	}

	function valueOf(name, fallback, formName) {
		var item = field(name, formName);
		return item ? item.value : fallback || "";
	}

	function setValue(name, value, formName) {
		var item = field(name, formName);
		if (item) {
			item.value = value == null ? "" : String(value);
		}
	}

	function selectedOption(select) {
		return select && select.options && select.selectedIndex >= 0 ? select.options[select.selectedIndex] : null;
	}

	function selectedValue(select) {
		var option = selectedOption(select);
		return option ? option.value : select && select.value || "";
	}

	function selectedText(select) {
		var option = selectedOption(select);
		return option ? option.text : "";
	}

	function toNumber(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function formatNumber(value, decimals) {
		return toNumber(value).toFixed(decimals == null ? 2 : decimals);
	}

	function dialogId() {
		var match = String(window.location.search || "").match(/[?&]__itmsDialogId=([^&]+)/);
		return match ? decodeURIComponent(match[1]) : "";
	}

	function ensureDialogArgs() {
		var id = dialogId();
		if (!window.dialogArguments && id && window.opener && window.opener.__itmsDialogArgs && id in window.opener.__itmsDialogArgs) {
			window.dialogArguments = window.opener.__itmsDialogArgs[id];
		}
		return window.dialogArguments;
	}

	function returnValue(value) {
		var id;
		window.returnValue = value;
		window.returnvalue = value;
		if (window.ITMSModalReturnCompat && window.ITMSModalReturnCompat.returnValue) {
			window.ITMSModalReturnCompat.returnValue(value);
			return;
		}
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(value);
			return;
		}
		id = dialogId();
		if (id && window.opener && window.opener.ITMSModernCompat && window.opener.ITMSModernCompat._receiveDialogValue) {
			window.opener.ITMSModernCompat._receiveDialogValue(id, value);
		}
	}

	function returnAndClose(value) {
		sent = true;
		returnValue(value);
		window.close();
	}

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		} else if (window.ITMSModernCompat && window.ITMSModernCompat.upgradeXmlIslands) {
			window.ITMSModernCompat.upgradeXmlIslands(document);
		}
	}

	function xmlObject(nameOrObject) {
		if (typeof nameOrObject !== "string") {
			return nameOrObject;
		}
		ensureCompat();
		return window[nameOrObject] || document[nameOrObject] || byId(nameOrObject) || null;
	}

	function xmlDocument(nameOrObject) {
		var object = xmlObject(nameOrObject);
		if (object && object.XMLDocument) {
			return object.XMLDocument;
		}
		if (object && object._doc) {
			return object._doc;
		}
		if (object && object.nodeType === 9) {
			return object;
		}
		if (object && object.documentElement) {
			return object;
		}
		return null;
	}

	function xmlRoot(nameOrObject) {
		var object = xmlObject(nameOrObject);
		if (object && object.documentElement) {
			return object.documentElement;
		}
		if (object && object.XMLDocument && object.XMLDocument.documentElement) {
			return object.XMLDocument.documentElement;
		}
		if (object && object._doc && object._doc.documentElement) {
			return object._doc.documentElement;
		}
		return null;
	}

	function createXmlElement(xmlName, nodeName) {
		var object = xmlObject(xmlName);
		var doc = xmlDocument(xmlName);
		if (object && typeof object.createElement === "function") {
			return object.createElement(nodeName);
		}
		if (doc && typeof doc.createElement === "function") {
			return doc.createElement(nodeName);
		}
		return document.implementation.createDocument("", "", null).createElement(nodeName);
	}

	function serializeXml(nameOrObject) {
		var doc = xmlDocument(nameOrObject);
		var root = xmlRoot(nameOrObject);
		if (doc) {
			return new XMLSerializer().serializeToString(doc);
		}
		return root ? new XMLSerializer().serializeToString(root) : "";
	}

	function childElements(node, name) {
		var wanted = name && String(name).toLowerCase();
		return Array.prototype.slice.call(node && node.childNodes || []).filter(function (child) {
			return child.nodeType === 1 && (!wanted || String(child.nodeName).toLowerCase() === wanted);
		});
	}

	function firstElement(node, name) {
		return childElements(node, name)[0] || null;
	}

	function getAttribute(node, name) {
		return node && node.getAttribute ? node.getAttribute(name) || "" : "";
	}

	function setAttribute(node, name, value) {
		if (node && node.setAttribute) {
			node.setAttribute(name, value == null ? "" : String(value));
		}
	}

	function clearChildren(node) {
		while (node && node.firstChild) {
			node.removeChild(node.firstChild);
		}
	}

	function installEscape(cancelFunction) {
		document.addEventListener("keydown", function (event) {
			if (event.key === "Escape" || event.keyCode === 27) {
				cancelFunction();
			}
		});
	}

	function installSimpleUnload(valueProvider) {
		window.window_onunload = function () {
			if (!sent) {
				returnValue(typeof valueProvider === "function" ? valueProvider() : valueProvider);
			}
		};
		window.addEventListener("beforeunload", window.window_onunload);
	}

	function openModernDialog(url, args, features, callback) {
		var popup;
		var timer;
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback);
			return;
		}
		popup = window.open(url, "_blank", "width=600,height=440,resizable=no,status=no,scrollbars=yes");
		if (!popup) {
			alert("Popup was blocked. Please allow popups for this site and try again.");
			return;
		}
		timer = window.setInterval(function () {
			if (popup.closed) {
				window.clearInterval(timer);
				if (typeof callback === "function") {
					callback();
				}
			}
		}, 500);
	}

	function installUnitSelection() {
		window.CheckSubmit = function () {
			var select = field("selUnit");
			if (!select || !select.value) {
				alert("Select Unit");
				if (select && select.focus) {
					select.focus();
				}
				return false;
			}
			returnAndClose(select.value + ":" + selectedText(select));
			return true;
		};
	}

	function installPartyDeleteSelection() {
		window.SelType = function () {
			var selectedUnits = [];
			var radios = fields("optAgsel");
			var unitSelect = field("selUnitId");
			var selectedRadio = radios.filter(function (radio) { return radio.checked; })[0];
			if (unitSelect && unitSelect.options) {
				Array.prototype.forEach.call(unitSelect.options, function (option) {
					if (option.selected) {
						selectedUnits.push(option.value);
					}
				});
			}
			if (selectedRadio && selectedRadio.value === "S" && selectedUnits.length === 0) {
				alert("Select atleast One Unit ");
				return false;
			}
			if (selectedRadio) {
				returnAndClose(selectedRadio.value === "A" ? selectedRadio.value : selectedUnits.join(","));
			}
			return true;
		};
	}

	function installPreferencePopup() {
		window.PageSubmit = function () {
			var frm = form("formname");
			if (!frm) {
				return false;
			}
			if (field("B2")) {
				field("B2").disabled = true;
			}
			if (field("B3")) {
				field("B3").disabled = true;
			}
			frm.action = "ParPrefPopupUpdate.asp";
			frm.submit();
			return true;
		};
		installSimpleUnload("Close");
	}

	function loadXmlIntoIsland(xmlName, xmlText) {
		var object = xmlObject(xmlName);
		if (object && typeof object.loadXML === "function") {
			object.loadXML(xmlText || "<Root/>");
		} else if (object) {
			object._doc = new DOMParser().parseFromString(xmlText || "<Root/>", "text/xml");
		}
		return xmlRoot(xmlName);
	}

	function installAgentSelect() {
		window.DisplayAgent = function () {
			var frm = form("formname");
			var unit = field("selUnitId");
			var agentType = field("selAgentType");
			var fromBox = field("selFrombox");
			var toBox = field("selTobox");
			var xhr;
			var root;
			if (!frm || !unit || !agentType || !fromBox || !toBox) {
				return false;
			}
			if (unit.selectedIndex === 0) {
				alert("Select Unit");
				unit.focus();
				return false;
			}
			fromBox.options.length = 0;
			toBox.options.length = 0;
			if (!valueOf("hPartyCode") || unit.selectedIndex <= 0 || agentType.selectedIndex <= 0) {
				return false;
			}
			xhr = new XMLHttpRequest();
			xhr.open("GET", "XMLPartyAgent.asp?Flag=A&AgentType=" + encodeURIComponent(agentType.value) + "&orgID=" + encodeURIComponent(selectedValue(unit)) + "&PartyCode=" + encodeURIComponent(valueOf("hPartyCode")), false);
			xhr.send(null);
			if (xhr.responseXML && xhr.responseXML.documentElement) {
				root = xhr.responseXML.documentElement;
				loadXmlIntoIsland("OutData", new XMLSerializer().serializeToString(xhr.responseXML));
			} else if (trim(xhr.responseText)) {
				root = loadXmlIntoIsland("OutData", xhr.responseText);
			}
			if (!root) {
				if (trim(xhr.responseText)) {
					alert(xhr.responseText);
				}
				return false;
			}
			childElements(root).forEach(function (node) {
				var target = getAttribute(node, "Selected") === "N" ? fromBox : toBox;
				target.options[target.options.length] = new Option(trim(node.textContent), getAttribute(node, "AgentId"));
			});
			return true;
		};
		installSimpleUnload("Close");
	}

	function selectedChecklistItems() {
		return fields("chkSelVal").filter(function (checkbox) {
			return checkbox.checked;
		});
	}

	function scheduleTarget() {
		return field("SelName2") || field("selName2");
	}

	function scheduleSource() {
		return field("SelName1") || field("selName1");
	}

	function scheduleOptions() {
		var target = scheduleTarget();
		return Array.prototype.slice.call(target && target.options || []);
	}

	function installScheduleSelection() {
		window.DisplayVal = function () {
			var unit = field("selUnitId");
			var frm = form("formname");
			if (unit && unit.selectedIndex !== 0 && frm) {
				frm.action = "SchSetupNew.asp";
				frm.submit();
			}
		};
		window.SubmitFun = function () {
			var options = scheduleOptions();
			var source = scheduleSource();
			var values;
			var names;
			var payload;
			var root;
			var doc;
			var xhr;
			if (options.length === 0) {
				alert("Select Name");
				if (source && source.focus) {
					source.focus();
				}
				return false;
			}
			if (config.mode === "returnList") {
				names = options.map(function (option) {
					return option.text;
				});
				values = options.map(function (option) {
					return option.value;
				});
				payload = names.join(":") + "~~" + values.join(",") + "~~" + options.length + "~~" + valueOf("hPass");
				setValue("hAcclist", payload);
				returnAndClose(payload);
				return true;
			}
			doc = document.implementation.createDocument("", "Root", null);
			root = doc.documentElement;
			options.forEach(function (option) {
				var parts = String(option.value || "").split("-");
				var node = doc.createElement("SchDetails");
				node.setAttribute("Description", option.text);
				node.setAttribute("ScheduleID", parts[0] || "");
				node.setAttribute("ScheduleSubID", parts[1] || "");
				node.setAttribute("ScheduleSubSubID", parts[2] || "");
				root.appendChild(node);
			});
			xhr = new XMLHttpRequest();
			xhr.open("POST", config.saveUrl || "XMLSave.asp?Name=SchedBSBrkSubHeads&Mod=Acc", false);
			xhr.send(new XMLSerializer().serializeToString(doc));
			if (trim(xhr.responseText) !== "") {
				alert(xhr.responseText);
			} else {
				returnAndClose(config.returnValue || "Y");
			}
			return true;
		};
		window.window_onunload = function () {
			var existing = valueOf("hAcclist");
			if (!sent && config.mode === "returnList") {
				if (!existing && scheduleOptions().length) {
					window.SubmitFun();
				} else {
					returnValue(existing);
				}
			}
		};
	}

	function installScheduleSetupCaller() {
		function submitTo(page) {
			var frm = form("formname");
			if (frm) {
				if (page) {
					frm.action = page;
				}
				frm.submit();
			}
		}

		function currentPage() {
			return config.page || window.location.pathname.split("/").pop();
		}

		function reloadPage() {
			submitTo(config.reloadPage || currentPage());
		}

		function setFieldIfPresent(names, value) {
			names.forEach(function (name) {
				if (field(name)) {
					setValue(name, value);
				}
			});
		}

		function setHtmlIfPresent(ids, html) {
			ids.forEach(function (id) {
				var item = byId(id);
				if (item) {
					item.innerHTML = html;
				}
			});
		}

		function openPagedAccountHead(url, callback) {
			openModernDialog(url, "", "dialogHeight:520px;dialogWidth:520px;center:Yes;help:No;resizable:No;status:No", function (value) {
				var parts = String(value || "").split(":");
				if (!trim(value)) {
					return;
				}
				if (parts.length <= 1) {
					openPagedAccountHead("ChgAccHeadName.asp?" + value, callback);
					return;
				}
				callback(parts, value);
			});
		}

		function addPopupUrl() {
			var url = config.addPopup || "AddSchedSubHeads.asp";
			var params = [
				"sUnit=" + encodeURIComponent(valueOf("hOrgID")),
				"sSchName=" + encodeURIComponent(valueOf("selSch")),
				"InsDate=" + encodeURIComponent(valueOf("selForMonth"))
			];
			if (config.addIncludesCat !== false) {
				params.splice(2, 0, "sCatCode=" + encodeURIComponent(config.addCatCode != null ? config.addCatCode : valueOf("hCatCode")));
			}
			return url + "?" + params.join("&");
		}

		function selectedScheduleValid() {
			var schedule = field("selSch");
			if (schedule && schedule.selectedIndex === 0) {
				alert("Select Schedule");
				schedule.focus();
				return false;
			}
			return true;
		}

		window.ViewSchedule = function (callFrom) {
			var url = "View_Schedule.asp?ForTheDate=" + encodeURIComponent(valueOf("selForMonth")) +
				"&OrgID=" + encodeURIComponent(valueOf("hOrgID")) +
				"&CallFrom=" + encodeURIComponent(callFrom || "");
			window.open(url, "", "height=540,width=795,toolbar=no,titlebar=no,location=no,directories=no,status=no,personalbar=no,menubar=no,scrollbars=no,resizable=yes,top=0,left=0");
			return true;
		};
		window.ViewPL = function () {
			var url = "View_ProfitAndLoss_Acc.asp?ForTheDate=" + encodeURIComponent(valueOf("selForMonth")) +
				"&OrgID=" + encodeURIComponent(valueOf("hOrgID"));
			window.open(url, "", "height=540,width=795,toolbar=no,titlebar=no,location=no,directories=no,status=no,personalbar=no,menubar=no,scrollbars=no,resizable=yes,top=0,left=0");
			return true;
		};
		window.CheckVal = function (item) {
			if (item && String(item.value || "").length > 100) {
				alert("Account Head Name Should be Less than 100 Characters ");
				item.focus();
				return false;
			}
			if (item && trim(item.value).length === 0) {
				alert("Account Head Name Should be blank ");
				item.focus();
				return false;
			}
			return true;
		};
		window.DisplayVal = function () {
			reloadPage();
			return true;
		};
		window.CheckSubmit = function () {
			if (!selectedScheduleValid()) {
				return false;
			}
			submitTo(config.updateAction || form("formname").action);
			return true;
		};
		window.popAccList = function (passCtr, entNo, entType) {
			setValue("hEntNo", entNo);
			setFieldIfPresent(["hType", "htype"], entType || "");
			openPagedAccountHead("ChgAccHeadName.asp?orgId=" + encodeURIComponent(valueOf("hOrgID")), function (parts) {
				var accountNo = parts[0] || "";
				var accountName = parts[1] || "";
				setValue("hAccHead" + entNo, accountNo);
				setValue("hAccHeadNo", accountNo);
				setValue("hAccHeadName", accountName);
				setHtmlIfPresent([
					"spAccHead" + passCtr + entNo,
					"spAccHead" + passCtr + " " + entNo
				], accountName);
			});
			return true;
		};
		window.SelAccList = function (passCtr, catCode, scheduleId, scheduleSubSubId, entType, shdType) {
			var scheduleKey = catCode + ":" + scheduleId + ":" + scheduleSubSubId;
			setValue("hSchdid", scheduleKey);
			setFieldIfPresent(["hType", "htype"], entType || "");
			openModernDialog("SelAccHeadName.asp?sTemp=" + encodeURIComponent(scheduleKey), "", "dialogHeight:480px;dialogWidth:700px;center:Yes;help:No;resizable:No;status:No", function (value) {
				var parts = String(value || "").split("~~");
				var accName = parts[0] || "";
				var accCode = parts[1] || "";
				var counter = parts[2] || "";
				var newSchSubId = parts[3] || "";
				var compactId = "spSchHead" + passCtr + scheduleId;
				var fullId = compactId + scheduleSubSubId;
				if (!trim(value)) {
					return;
				}
				if (String(shdType || "").toUpperCase() === "A") {
					setHtmlIfPresent([fullId, "spSchHead" + passCtr + " " + scheduleId + " " + scheduleSubSubId], accName);
					setValue("hSchHead" + passCtr + scheduleId + scheduleSubSubId, value);
				} else {
					setHtmlIfPresent([compactId, fullId, "spSchHead" + passCtr + " " + scheduleId], accName);
					setFieldIfPresent([
						"hSchHead" + passCtr + scheduleId,
						"hSchHead" + passCtr + scheduleId + scheduleSubSubId
					], value);
					setValue("hAccCode", valueOf("hAccCode") + "*" + accCode + ":" + counter + ":" + newSchSubId);
				}
			});
			return true;
		};
		window.SchdSetupPopUp = function (callFrom) {
			var orgId = valueOf("hOrgID");
			var scheduleId = valueOf("selSch");
			var insDate = valueOf("selForMonth");
			if (String(orgId) === "0") {
				alert("Select Organization And Continue...!");
				return false;
			}
			openModernDialog("BalScheduleSetUp.asp?sUnit=" + encodeURIComponent(orgId) + "&sSchID=" + encodeURIComponent(scheduleId) + "&InsDate=" + encodeURIComponent(insDate) + "&CallFrom=" + encodeURIComponent(callFrom || ""), "A", "dialogHeight:250px;dialogWidth:620px;center:Yes;help:No;resizable:No;status:No", function (value) {
				if (String(value) === "Y") {
					reloadPage();
				}
			});
			return true;
		};
		window.AddPopUp = function () {
			openModernDialog(addPopupUrl(), "A", config.addFeatures || "dialogHeight:330px;dialogWidth:760px;center:Yes;help:No;resizable:No;status:No", function (value) {
				if (String(value) === "Y") {
					reloadPage();
				}
			});
			return true;
		};
		window.Del = function () {
			var root = xmlRoot("TempData");
			var node;
			var response;
			if (!root) {
				loadXmlIntoIsland("TempData", "<Root/>");
				root = xmlRoot("TempData");
			}
			clearChildren(root);
			node = createXmlElement("TempData", "Schedule");
			node.setAttribute("id", "5");
			node.setAttribute("sOrgID", valueOf("hOrgID"));
			node.setAttribute("sschedno", valueOf("selSch"));
			root.appendChild(node);
			response = postText("XMLShdDelete_Update.asp?Name=SchdDelete&Mod=Acc", serializeXml("TempData"));
			if (trim(response) !== "") {
				alert(response);
			} else {
				submitTo(config.deleteReloadPage || "PLSetup.asp");
			}
			return true;
		};
		window.Checkreturn = function () {
			return true;
		};
	}

	function installBalanceScheduleSetup() {
		window.Window = window;

		function resetRoot() {
			var current = xmlRoot("TempData");
			if (!current) {
				loadXmlIntoIsland("TempData", "<Root/>");
				current = xmlRoot("TempData");
			}
			clearChildren(current);
			return current;
		}

		function numberCheck(name, message) {
			var item = field(name);
			if (item && isNaN(Number(item.value))) {
				alert(message);
				item.value = "";
				item.focus();
				return false;
			}
			return true;
		}

		window.CheckVal = function () {
			if (!numberCheck("txtno", "Enter Numerals Only")) {
				return false;
			}
			if (trim(valueOf("txtSchdHead")) !== "" && !isNaN(Number(valueOf("txtSchdHead")))) {
				alert("Enter Schedule Heading");
				setValue("txtSchdHead", "");
				if (field("txtSchdHead")) {
					field("txtSchdHead").focus();
				}
				return false;
			}
			return true;
		};
		window.ChkVal = function () {
			return numberCheck("txtSchdHiera", "Enter Numerals for Hierarchy");
		};
		window.LoadVal = function () {
			alert(valueOf("iSchId"));
			return true;
		};
		window.CheckSubmit = function () {
			var scheduleRoot = resetRoot();
			var node;
			var response;
			if (trim(valueOf("txtno")) === "") {
				alert("Enter Schedule No");
				return false;
			}
			if (trim(valueOf("txtSchdHead")) === "") {
				alert("Enter Schedule Name");
				return false;
			}
			if (trim(valueOf("txtSchdHiera")) === "") {
				alert("Enter Schedule Hierarchy");
				return false;
			}
			if (trim(valueOf("selApp")) === "") {
				alert("Select Applicable for");
				return false;
			}
			if (trim(valueOf("FinYear")) === "") {
				alert("Select Year");
				return false;
			}
			node = createXmlElement("TempData", "Schedule");
			node.setAttribute("PrevSchNo", valueOf("iSchId"));
			node.setAttribute("SchedNo", valueOf("txtno"));
			node.setAttribute("OrgID", valueOf("hOrgId"));
			node.setAttribute("SchedHead", valueOf("txtSchdHead"));
			node.setAttribute("SchedHiera", valueOf("txtSchdHiera"));
			node.setAttribute("SchedApp", valueOf("selApp"));
			node.setAttribute("SchedYear", valueOf("FinYear"));
			node.setAttribute("InsDate", valueOf("hInsDate"));
			scheduleRoot.appendChild(node);
			response = postText("SchedInsert.asp?Name=Sched&Mod=Acc", serializeXml("TempData"));
			if (trim(response) !== "") {
				alert(response);
			} else {
				returnAndClose("Y");
			}
			return true;
		};
		window.Del = function () {
			var scheduleRoot = resetRoot();
			var node = createXmlElement("TempData", "Schedule");
			var response;
			node.setAttribute("sNo", valueOf("txtno"));
			node.setAttribute("sOrgId", valueOf("hOrgId"));
			node.setAttribute("sFinYr", valueOf("FinYear"));
			node.setAttribute("id", "3");
			scheduleRoot.appendChild(node);
			response = postText("XMLShdDelete_Update.asp?Name=SchdDelete&Mod=Acc", serializeXml("TempData"));
			if (trim(response) !== "") {
				alert(response);
			} else {
				returnAndClose("Y");
			}
			return true;
		};
	}

	function installScheduleSubHeadsPopup() {
		function selectedRadioIndex(name) {
			var list = fields(name);
			var index;
			for (index = 0; index < list.length; index += 1) {
				if (list[index].checked) {
					return index;
				}
			}
			return -1;
		}

		function setRadioIndex(name, index) {
			var list = fields(name);
			if (list[index]) {
				list[index].checked = true;
			}
		}

		function selectedTextByName(name) {
			return selectedText(field(name));
		}

		function resetSelect(name, firstText, firstValue) {
			var select = field(name);
			if (!select) {
				return null;
			}
			select.options.length = 0;
			select.options[0] = new Option(firstText, firstValue);
			return select;
		}

		function getXml(url) {
			var xhr = new XMLHttpRequest();
			var root;
			xhr.open("GET", url, false);
			xhr.send(null);
			if (xhr.responseXML && xhr.responseXML.documentElement) {
				loadXmlIntoIsland("OutData", new XMLSerializer().serializeToString(xhr.responseXML));
			} else if (trim(xhr.responseText) !== "") {
				loadXmlIntoIsland("OutData", xhr.responseText);
			}
			root = xmlRoot("OutData");
			if (!root && trim(xhr.responseText)) {
				alert(xhr.responseText);
			}
			return root;
		}

		function modeType() {
			var index = selectedRadioIndex("optMode");
			return index === 0 ? "D" : index === 1 ? "A" : index === 2 ? "S" : "N";
		}

		function setMode(type) {
			if (type === "D") {
				setRadioIndex("optMode", 0);
			} else if (type === "A") {
				setRadioIndex("optMode", 1);
			} else if (type === "S") {
				setRadioIndex("optMode", 2);
			} else {
				setRadioIndex("optMode", 3);
			}
			window.ModeFun();
		}

		function setComputeMode(value) {
			setRadioIndex("optCompMode", value === "+" ? 0 : 1);
		}

		function levelOneParts() {
			return String(valueOf("selLevel1")).split("-");
		}

		function fillAccountFromServer(entryType, subId, subSubId) {
			var url = "AccSubAndSubSubID.asp?sOrgID=" + encodeURIComponent(valueOf("hOrgId")) +
				"&EntryType=" + encodeURIComponent(entryType || "") +
				"&sFinyr=" + encodeURIComponent(valueOf("sfinyr")) +
				"&sschedno=" + encodeURIComponent(valueOf("sschedno")) +
				"&SubID=" + encodeURIComponent(subId || "") +
				"&SubSubID=" + encodeURIComponent(subSubId || "") +
				"&id=3";
			var root = getXml(url);
			var node = childElements(root)[0];
			if (node) {
				setValue("hAccHead", attrAt(node, 0));
				setValue("txtAcHead", attrAt(node, 1));
			}
		}

		function openPagedAccountHead(url) {
			openModernDialog(url, "", "dialogHeight:520px;dialogWidth:520px;center:Yes;help:No;resizable:No;status:No", function (value) {
				var parts = String(value || "").split(":");
				if (!trim(value)) {
					return;
				}
				if (parts.length <= 1) {
					openPagedAccountHead("ChgAccHeadName.asp?" + value);
					return;
				}
				setValue("hAccHead", parts[0] || "");
				setValue("txtAcHead", parts[1] || "");
			});
		}

		function detailsNode() {
			var root = xmlRoot("XmlData");
			var node;
			if (!root) {
				loadXmlIntoIsland("XmlData", "<Root><Details OrgID=\"\" SchID=\"\" LevelID=\"\" Level1ID=\"\" Level2ID=\"\" Level1Name=\"\" Level2Name=\"\" ModeType=\"\" AccHead=\"\" AccHeadName=\"\" FinYear=\"\" ComputeMode=\"\" InsDate=\"" + valueOf("hInsDate") + "\" /></Root>");
				root = xmlRoot("XmlData");
			}
			node = firstElement(root, "Details");
			if (!node) {
				node = createXmlElement("XmlData", "Details");
				root.appendChild(node);
			}
			return node;
		}

		function tempRoot() {
			var root = xmlRoot("TempData");
			if (!root) {
				loadXmlIntoIsland("TempData", "<Root/>");
				root = xmlRoot("TempData");
			}
			clearChildren(root);
			return root;
		}

		function scheduleField() {
			return field("selSch") || field("SelSch");
		}

		function scheduleValue() {
			var item = scheduleField();
			return item ? item.value : "";
		}

		window.SetLevelFun = function () {
			var level1 = field("selLevel1");
			var level2;
			var parts;
			var root;
			if (!level1 || level1.value === "0") {
				alert("Select SubHeading And Proceed...!");
				return false;
			}
			if (field("txtLev1")) {
				field("txtLev1").disabled = false;
				field("txtLev1").value = level1.value !== "A" ? selectedText(level1) : "";
				field("txtLev1").size = field("txtLev1").value ? field("txtLev1").value.length + 7 : field("txtLev1").size;
			}
			if (field("txtLev2")) {
				field("txtLev2").disabled = false;
			}
			level2 = resetSelect("selLevel2", "Add New", "A");
			if (level1.value !== "A") {
				parts = levelOneParts();
				setComputeMode(parts[3]);
				if (selectedRadioIndex("optLevel") === 1) {
					root = getXml("AccSubAndSubSubID.asp?sOrgID=" + encodeURIComponent(valueOf("hOrgId")) +
						"&EntryType=" + encodeURIComponent(parts[2] || "") +
						"&sFinyr=" + encodeURIComponent(valueOf("sfinyr")) +
						"&sschedno=" + encodeURIComponent(valueOf("sschedno")) +
						"&SubID=" + encodeURIComponent(parts[0] || "") +
						"&SubSubID=" + encodeURIComponent(parts[1] || "") +
						"&id=2");
					childElements(root).forEach(function (node) {
						level2.options[level2.options.length] = new Option(attrAt(node, 1), attrAt(node, 0));
					});
				}
				setMode(parts[2]);
				if (parts[2] === "A") {
					fillAccountFromServer(parts[2], parts[0], parts[1]);
				}
			}
			return true;
		};

		window.LevelFun = function () {
			var levelTwo = selectedRadioIndex("optLevel") !== 0;
			setValue("txtLev2", levelTwo ? valueOf("txtLev2") : "");
			if (field("txtLev2")) {
				field("txtLev2").disabled = !levelTwo;
			}
			if (field("selLevel2")) {
				field("selLevel2").disabled = !levelTwo;
			}
			if (field("FinYear")) {
				field("FinYear").disabled = !levelTwo;
			}
			return true;
		};

		window.setlevelfun1 = function () {
			var parts;
			var level1;
			if (valueOf("selLevel2") !== "A") {
				parts = String(valueOf("selLevel2")).split("-");
				level1 = levelOneParts();
				setValue("txtLev2", selectedTextByName("selLevel2"));
				if (field("txtLev2")) {
					field("txtLev2").disabled = false;
				}
				setMode(parts[1]);
				if (parts[1] === "A") {
					setValue("hAccHead", parts[2] || "");
					fillAccountFromServer(parts[1], level1[0], parts[0]);
				}
			} else {
				setValue("txtLev2", "");
			}
			return true;
		};

		window.ModeFun = function () {
			var isAccountMode = selectedRadioIndex("optMode") === 1;
			if (field("ButAcHead")) {
				field("ButAcHead").disabled = !isAccountMode;
			}
			if (!isAccountMode) {
				setValue("txtAcHead", "");
			}
			return true;
		};
		window.AccHeadClck = function () {
			if (field("txtAcHead")) {
				field("txtAcHead").disabled = false;
			}
			return true;
		};
		window.popAccList = function () {
			openPagedAccountHead("ChgAccHeadName.asp?orgId=" + encodeURIComponent(valueOf("hOrgId")));
			return true;
		};
		window.CheckVal = function () {
			if (String(scheduleValue()) === "0") {
				alert("Select Schedule ");
				if (scheduleField() && scheduleField().focus) {
					scheduleField().focus();
				}
				return false;
			}
			if (field("selLevel1") && field("selLevel1").selectedIndex === 0 && trim(valueOf("txtLev1")) === "") {
				alert("Select Level 1 ");
				field("selLevel1").focus();
				return false;
			}
			if (trim(valueOf("txtLev2")) === "" && field("selLevel2") && field("selLevel2").selectedIndex === 0 && !field("selLevel2").disabled) {
				alert("Select Level 2 ");
				field("selLevel2").focus();
				return false;
			}
			if (selectedRadioIndex("optMode") === 1 && String(valueOf("hAccHead")) === "0") {
				alert("Select Account Head ");
				return false;
			}
			return true;
		};
		window.SaveXML = function () {
			var node;
			var response;
			if (!window.CheckVal()) {
				return false;
			}
			node = detailsNode();
			setAttribute(node, "OrgID", valueOf("hOrgId"));
			setAttribute(node, "SchID", scheduleValue());
			setAttribute(node, "LevelID", selectedRadioIndex("optLevel") === 0 ? "0" : "1");
			setAttribute(node, "Level1ID", valueOf("selLevel1"));
			setAttribute(node, "Level2ID", valueOf("selLevel2"));
			setAttribute(node, "Level1Name", valueOf("txtLev1"));
			setAttribute(node, "Level2Name", valueOf("txtLev2"));
			setAttribute(node, "ModeType", modeType());
			setAttribute(node, "AccHead", valueOf("hAccHead"));
			if (modeType() === "A") {
				setAttribute(node, "AccHeadName", valueOf("txtAcHead"));
			}
			setAttribute(node, "FinYear", valueOf("FinYear"));
			setAttribute(node, "ComputeMode", selectedRadioIndex("optCompMode") === 0 ? "+" : "-");
			response = postText("XMLSchHeadSave.asp?Name=SchedSubHeads&Mod=Acc", serializeXml("XmlData"));
			if (trim(response) !== "") {
				alert(response);
			} else {
				returnAndClose("Y");
			}
			return true;
		};
		window.Del = function () {
			var level1 = levelOneParts();
			var level2 = String(valueOf("selLevel2")).split("-");
			var subId = level1[0] || "";
			var subSubId = level1[1] || "";
			var root = tempRoot();
			var node = createXmlElement("TempData", "Schedule");
			var response;
			if (field("selLevel2") && field("selLevel2").length > 1 && valueOf("selLevel2") !== "A") {
				subSubId = level2[0] || "";
			}
			node.setAttribute("id", "1");
			node.setAttribute("sschedno", valueOf("sschedno"));
			node.setAttribute("SubID", subId);
			node.setAttribute("SubSubID", subSubId);
			node.setAttribute("sOrgID", valueOf("hOrgId"));
			node.setAttribute("sFinyr", valueOf("sfinyr"));
			root.appendChild(node);
			response = postText("XMLShdDelete_Update.asp?Name=SchdDelete&Mod=Acc", serializeXml("TempData"));
			if (trim(response) !== "") {
				alert(response);
			} else {
				returnAndClose("Y");
			}
			return true;
		};
		window.EditFields = function () {
			if (valueOf("selLevel1") !== "A") {
				if (field("txtLev1")) {
					field("txtLev1").disabled = false;
					field("txtLev1").value = selectedTextByName("selLevel1");
					field("txtLev1").size = trim(field("txtLev1").value).length + 7;
				}
			}
			if (valueOf("selLevel2") !== "A" && field("txtLev2")) {
				field("txtLev2").disabled = false;
				field("txtLev2").value = selectedTextByName("selLevel2");
			}
			if (field("btnEdit")) {
				field("btnEdit").disabled = true;
			}
			return true;
		};
	}

	function installPlBsScheduleSubHeadsPopup() {
		var kind = String(config.kind || "PL").toUpperCase() === "BS" ? "BS" : "PL";
		var saveUrl = kind === "BS" ? "XMLSchBSHeadSave.asp?Name=SchedBSSubHeads&Mod=Acc" : "XMLSchPLHeadSave.asp?Name=SchedPLSubHeads&Mod=Acc";
		var breakupSaveUrl = kind === "BS" ? "XMLSave.asp?Name=SchedBSBrkSubHeads&Mod=Acc" : "XMLSave.asp?Name=SchedPLBrkSubHeads&Mod=Acc";
		var deleteId = kind === "BS" ? "6" : "4";

		function selectedRadioIndex(name) {
			var list = fields(name);
			var index;
			for (index = 0; index < list.length; index += 1) {
				if (list[index].checked) {
					return index;
				}
			}
			return -1;
		}

		function setRadioIndex(name, index) {
			var list = fields(name);
			if (list[index]) {
				list[index].checked = true;
			}
		}

		function splitCsv(value) {
			return String(value || "").split(",");
		}

		function modeType() {
			var index = selectedRadioIndex("optMode");
			return index === 0 ? "D" : index === 1 ? "A" : index === 2 ? "S" : "N";
		}

		function setMode(type) {
			if (type === "D") {
				setRadioIndex("optMode", 0);
			} else if (type === "A") {
				setRadioIndex("optMode", 1);
			} else if (type === "S") {
				setRadioIndex("optMode", 2);
			} else {
				setRadioIndex("optMode", 3);
			}
			window.ModeFun();
		}

		function setComputeMode(value) {
			setRadioIndex("optCompMode", value === "+" ? 0 : 1);
		}

		function resetSelect(name, firstText, firstValue) {
			var select = field(name);
			if (!select) {
				return null;
			}
			select.options.length = 0;
			select.options[0] = new Option(firstText, firstValue);
			return select;
		}

		function getXml(url) {
			var xhr = new XMLHttpRequest();
			xhr.open("GET", url, false);
			xhr.send(null);
			if (xhr.responseXML && xhr.responseXML.documentElement) {
				loadXmlIntoIsland("OutData", new XMLSerializer().serializeToString(xhr.responseXML));
			} else if (trim(xhr.responseText) !== "") {
				loadXmlIntoIsland("OutData", xhr.responseText);
			}
			if (!xmlRoot("OutData") && trim(xhr.responseText)) {
				alert(xhr.responseText);
			}
			return xmlRoot("OutData");
		}

		function fillAccountByCode(accountCode) {
			var root;
			var node;
			if (!trim(accountCode)) {
				return;
			}
			root = getXml("AccSubAndSubSubID.asp?sOrgID=" + encodeURIComponent(valueOf("hOrgId")) +
				"&sFinyr=" + encodeURIComponent(valueOf("sfinyr")) +
				"&AcCode=" + encodeURIComponent(accountCode) +
				"&id=7");
			node = childElements(root)[0];
			if (node) {
				setValue("hAccHead", attrAt(node, 0));
				setValue("txtAcHead", attrAt(node, 1));
			}
		}

		function postScheduleBreakup(subId, subSubId) {
			var xhr = new XMLHttpRequest();
			var xmlText;
			xhr.open("GET", "AccSubAndSubSubID.asp?sOrgID=" + encodeURIComponent(valueOf("hOrgId")) +
				"&sFinyr=" + encodeURIComponent(valueOf("sfinyr")) +
				"&sschedno=" + encodeURIComponent(valueOf("selSch")) +
				"&SubID=" + encodeURIComponent(subId || "") +
				"&SubSubID=" + encodeURIComponent(subSubId || "") +
				"&id=6", false);
			xhr.send(null);
			xmlText = xhr.responseXML && xhr.responseXML.documentElement ? new XMLSerializer().serializeToString(xhr.responseXML) : xhr.responseText;
			if (trim(xmlText) !== "") {
				postText(breakupSaveUrl, xmlText);
			}
		}

		function openPagedAccountHead(url) {
			openModernDialog(url, "", "dialogHeight:520px;dialogWidth:520px;center:Yes;help:No;resizable:No;status:No", function (value) {
				var parts = String(value || "").split(":");
				if (!trim(value)) {
					return;
				}
				if (parts.length <= 1) {
					openPagedAccountHead("ChgAccHeadName.asp?" + value);
					return;
				}
				setValue("hAccHead", parts[0] || "");
				setValue("txtAcHead", parts[1] || "");
			});
		}

		function selectionPopupUrl() {
			var headParam = kind + "Head";
			var subParam = kind + "SubHd";
			var subSubParam = kind === "BS" ? "BSSubSubHd" : "PlSubSubHd";
			var page = kind === "BS" ? "SelAccBSHeadName.asp" : "SelAccPLHeadName.asp";
			var subSub = selectedRadioIndex("optLevel") === 1 ? valueOf("selLevel2") : "0";
			return page + "?orgId=" + encodeURIComponent(valueOf("hOrgId")) +
				"&" + headParam + "=" + encodeURIComponent(valueOf("selSch")) +
				"&" + subParam + "=" + encodeURIComponent(valueOf("selLevel1")) +
				"&" + subSubParam + "=" + encodeURIComponent(subSub);
		}

		function detailsNode() {
			var root = xmlRoot("XmlData");
			var node;
			if (!root) {
				loadXmlIntoIsland("XmlData", "<Root><Details OrgID=\"\" SchName=\"\" SchID=\"\" LevelID=\"\" Level1ID=\"\" Level2ID=\"\" Level1Name=\"\" Level2Name=\"\" ModeType=\"\" AccHead=\"\" AccHeadName=\"\" FinYear=\"\" ComputeMode=\"\" Hierachy=\"\" InsDate=\"\" /></Root>");
				root = xmlRoot("XmlData");
			}
			node = firstElement(root, "Details");
			if (!node) {
				node = createXmlElement("XmlData", "Details");
				root.appendChild(node);
			}
			return node;
		}

		function tempRoot() {
			var root = xmlRoot("TempData");
			if (!root) {
				loadXmlIntoIsland("TempData", "<Root/>");
				root = xmlRoot("TempData");
			}
			clearChildren(root);
			return root;
		}

		window.SelHead = function () {
			var schedule = field("selSch");
			var level1;
			var root;
			if (!schedule) {
				return false;
			}
			if (schedule.value === "A") {
				if (field("txtLev")) {
					field("txtLev").disabled = false;
				}
				setValue("txtLev", "");
				level1 = resetSelect("selLevel1", "AddNew", "A");
				if (field("txtLev1")) {
					field("txtLev1").disabled = false;
				}
			} else {
				setValue("txtLev", selectedText(schedule));
				level1 = resetSelect("selLevel1", "Add New", "A");
				root = getXml("AccSubAndSubSubID.asp?sOrgID=" + encodeURIComponent(valueOf("hOrgId")) +
					"&sFinyr=" + encodeURIComponent(valueOf("sfinyr")) +
					"&sschedno=" + encodeURIComponent(schedule.value) +
					"&id=4");
				childElements(root).forEach(function (node) {
					level1.options[level1.options.length] = new Option(attrAt(node, 1), attrAt(node, 0));
				});
			}
			return true;
		};

		window.SetLevelFun = function () {
			var level1 = field("selLevel1");
			var level2;
			var value;
			var root;
			if (!level1) {
				return false;
			}
			if (field("txtLev1")) {
				field("txtLev1").disabled = false;
			}
			if (field("txtLev2")) {
				field("txtLev2").disabled = false;
			}
			if (level1.value === "A") {
				resetSelect("selLevel2", "AddNew", "A");
				return true;
			}
			value = splitCsv(level1.value);
			setValue("txtLev1", selectedText(level1));
			setValue("txtHierarchy", value[5] || "");
			if (value[2] === "A") {
				setMode("A");
				setValue("hAccHead", value[4] || "");
				fillAccountByCode(value[4]);
			} else if (value[2] === "S") {
				setMode("S");
				setValue("hAccHead", value[4] || "");
				postScheduleBreakup(value[0], value[1]);
			} else if (value[2] === "D") {
				setMode("D");
			} else {
				setMode("N");
			}
			setComputeMode(value[3]);
			level2 = resetSelect("selLevel2", "Add New", "A");
			root = getXml("AccSubAndSubSubID.asp?sOrgID=" + encodeURIComponent(valueOf("hOrgId")) +
				"&sFinyr=" + encodeURIComponent(valueOf("sfinyr")) +
				"&sschedno=" + encodeURIComponent(valueOf("selSch")) +
				"&SubID=" + encodeURIComponent(value[0] || "") +
				"&id=5");
			childElements(root).forEach(function (node) {
				level2.options[level2.options.length] = new Option(attrAt(node, 1), attrAt(node, 0));
			});
			return true;
		};

		window.LevelFun = function () {
			var levelTwo = selectedRadioIndex("optLevel") !== 0;
			if (!levelTwo) {
				setValue("txtLev2", "");
				setValue("txtAcHead", "");
			}
			if (field("txtLev2")) {
				field("txtLev2").disabled = !levelTwo;
			}
			if (field("txtAcHead")) {
				field("txtAcHead").disabled = !levelTwo;
			}
			if (field("selLevel2")) {
				field("selLevel2").disabled = !levelTwo;
			}
			if (field("FinYear")) {
				field("FinYear").disabled = !levelTwo;
			}
			return true;
		};

		window.setlevelfun1 = function () {
			var value;
			if (valueOf("selLevel2") === "A") {
				setValue("txtLev2", "");
				return true;
			}
			value = splitCsv(valueOf("selLevel2"));
			setValue("txtLev2", selectedText(field("selLevel2")));
			setValue("txtHierarchy", value[5] || "");
			if (field("txtLev2")) {
				field("txtLev2").disabled = false;
			}
			if (value[2] === "A") {
				setMode("A");
				setValue("hAccHead", value[4] || "");
				fillAccountByCode(value[4]);
			} else if (value[2] === "S") {
				setMode("S");
				setValue("hAccHead", value[4] || "");
				postScheduleBreakup(value[0], value[1]);
			} else if (value[2] === "D") {
				setMode("D");
			} else {
				setMode("N");
			}
			setComputeMode(value[3]);
			return true;
		};

		window.ModeFun = function () {
			var modeIndex = selectedRadioIndex("optMode");
			var selectable = modeIndex === 1 || modeIndex === 2;
			if (field("ButAcHead")) {
				field("ButAcHead").disabled = !selectable;
			}
			if (!selectable) {
				setValue("txtAcHead", "");
			}
			return true;
		};
		window.AccHeadClck = function () {
			if (field("txtAcHead")) {
				field("txtAcHead").disabled = false;
			}
			return true;
		};
		window.popAccList = function () {
			if (selectedRadioIndex("optMode") === 1) {
				openPagedAccountHead("ChgAccHeadName.asp?orgId=" + encodeURIComponent(valueOf("hOrgId")));
			} else if (selectedRadioIndex("optMode") === 2) {
				openModernDialog(selectionPopupUrl(), "A", "dialogHeight:460px;dialogWidth:550px;center:Yes;help:No;resizable:Yes;status:No", function () {});
			}
			return true;
		};
		window.CheckVal = function () {
			if (field("selSch") && field("selSch").selectedIndex === 0) {
				alert("Select Schedule ");
				field("selSch").focus();
				return false;
			}
			if (field("selLevel1") && field("selLevel1").selectedIndex === 0 && valueOf("selLevel1") !== "A") {
				alert("Select Level 1 ");
				field("selLevel1").focus();
				return false;
			}
			if (selectedRadioIndex("optLevel") === 1 && trim(valueOf("txtLev2")) === "") {
				alert("Select Level 2 ");
				if (field("selLevel2")) {
					field("selLevel2").focus();
				}
				return false;
			}
			if (selectedRadioIndex("optMode") === 1 && String(valueOf("hAccHead")) === "0") {
				alert("Select Account Head ");
				return false;
			}
			return true;
		};
		window.SaveXML = function () {
			var node;
			var response;
			if (valueOf("selSch") === "A" && (trim(valueOf("txtLev")) === "" || trim(valueOf("txtLev1")) === "")) {
				alert("Enter All The Information And Try To Save...!");
				return false;
			}
			if (!window.CheckVal()) {
				return false;
			}
			if (valueOf("selSch") === "A" && trim(valueOf("txtLev")) === "") {
				alert("Enter Schedule And Continue...!");
				return false;
			}
			node = detailsNode();
			setAttribute(node, "SchName", valueOf("txtLev"));
			setAttribute(node, "OrgID", valueOf("hOrgId"));
			setAttribute(node, "SchID", valueOf("selSch"));
			setAttribute(node, "LevelID", selectedRadioIndex("optLevel") === 0 ? "0" : "1");
			setAttribute(node, "Level1ID", valueOf("selLevel1"));
			setAttribute(node, "Level2ID", valueOf("selLevel2"));
			setAttribute(node, "Level1Name", valueOf("txtLev1"));
			setAttribute(node, "Level2Name", valueOf("txtLev2"));
			setAttribute(node, "ModeType", modeType());
			setAttribute(node, "AccHead", valueOf("hAccHead"));
			if (modeType() === "A") {
				setAttribute(node, "AccHeadName", valueOf("txtAcHead"));
			}
			setAttribute(node, "FinYear", valueOf("FinYear"));
			setAttribute(node, "ComputeMode", selectedRadioIndex("optCompMode") === 0 ? "+" : "-");
			setAttribute(node, "Hierachy", valueOf("txtHierarchy"));
			response = postText(saveUrl, serializeXml("XmlData"));
			if (trim(response) !== "") {
				alert(response);
			} else {
				returnAndClose("Y");
			}
			return true;
		};
		window.Del = function () {
			var level1 = splitCsv(valueOf("selLevel1"));
			var level2 = splitCsv(valueOf("selLevel2"));
			var subId = "0";
			var subSubId = "0";
			var root;
			var node;
			var response;
			if (valueOf("selSch") === "A") {
				alert("Select Any Value And Then Delete...!");
				return false;
			}
			if (selectedRadioIndex("optLevel") === 0 && valueOf("selLevel1") !== "A") {
				subId = level1[0] || "0";
			}
			if (selectedRadioIndex("optLevel") === 1 && valueOf("selLevel2") !== "A") {
				subId = level1[0] || "0";
				subSubId = level2[1] || "0";
			}
			root = tempRoot();
			node = createXmlElement("TempData", "Schedule");
			node.setAttribute("id", deleteId);
			node.setAttribute("sschedno", valueOf("selSch"));
			node.setAttribute("SubID", subId);
			node.setAttribute("SubSubID", subSubId);
			node.setAttribute("sOrgID", valueOf("hOrgId"));
			node.setAttribute("sFinyr", valueOf("sfinyr"));
			root.appendChild(node);
			response = postText("XMLShdDelete_Update.asp?Name=SchdDelete&Mod=Acc", serializeXml("TempData"));
			if (trim(response) !== "") {
				alert(response);
			} else {
				returnAndClose("Y");
			}
			return true;
		};
		window.DispName = function () {
			var schedule = field("selSch");
			if (schedule) {
				setValue("txtLev", selectedText(schedule));
			}
			return true;
		};
	}

	function installScheduleBreakupSubHeadsPopup() {
		function selectedRadioIndex(name) {
			var list = fields(name);
			var index;
			for (index = 0; index < list.length; index += 1) {
				if (list[index].checked) {
					return index;
				}
			}
			return -1;
		}

		function setRadioIndex(name, index) {
			var list = fields(name);
			if (list[index]) {
				list[index].checked = true;
			}
		}

		function splitDash(value) {
			return String(value || "").split("-");
		}

		function splitCsv(value) {
			return String(value || "").split(",");
		}

		function resetSelect(name, firstText, firstValue) {
			var select = field(name);
			if (!select) {
				return null;
			}
			select.options.length = 0;
			select.options[0] = new Option(firstText, firstValue);
			return select;
		}

		function getXml(url) {
			var xhr = new XMLHttpRequest();
			xhr.open("GET", url, false);
			xhr.send(null);
			if (xhr.responseXML && xhr.responseXML.documentElement) {
				loadXmlIntoIsland("OutData", new XMLSerializer().serializeToString(xhr.responseXML));
			} else if (trim(xhr.responseText) !== "") {
				loadXmlIntoIsland("OutData", xhr.responseText);
			}
			if (!xmlRoot("OutData") && trim(xhr.responseText)) {
				alert(xhr.responseText);
			}
			return xmlRoot("OutData");
		}

		function tempRoot() {
			var root = xmlRoot("TempData");
			if (!root) {
				loadXmlIntoIsland("TempData", "<Root/>");
				root = xmlRoot("TempData");
			}
			clearChildren(root);
			return root;
		}

		function appendOption(select, text, value) {
			if (select) {
				select.options[select.options.length] = new Option(text, value);
			}
		}

		function selectedTextByName(name) {
			return selectedText(field(name));
		}

		function openPagedAccountHead(url) {
			openModernDialog(url, "", "dialogHeight:520px;dialogWidth:520px;center:Yes;help:No;resizable:No;status:No", function (value) {
				var parts = String(value || "").split(":");
				if (!trim(value)) {
					return;
				}
				if (parts.length <= 1) {
					openPagedAccountHead("ChgAccHeadName.asp?" + value);
					return;
				}
				setValue("hAccHead", parts[0] || "");
				setValue("txtAcHead", parts[1] || "");
			});
		}

		window.SetLevelFun = function () {
			var level1 = field("sel1");
			var level2;
			var parts;
			var root;
			if (!level1) {
				return false;
			}
			if (level1.value === "0") {
				alert("Select Head Name And Proceed...");
			}
			if (field("txtLev1")) {
				field("txtLev1").disabled = !(level1.value === "A" || level1.selectedIndex === 0);
			}
			if (selectedRadioIndex("optLevel") !== 0) {
				level2 = resetSelect("sel2", "Add New", "A");
				if (field("txtLev2")) {
					field("txtLev2").disabled = false;
				}
				if (level1.value !== "A" && level1.selectedIndex !== 0) {
					parts = splitDash(level1.value);
					setValue("ShSubID", parts[0] || "");
					setValue("ShSubSubID", parts[1] || "");
					root = getXml("AccSubAndSubSubID.asp?sOrgID=" + encodeURIComponent(valueOf("sUnit")) +
						"&sShID=" + encodeURIComponent(valueOf("sschedno")) +
						"&sFinyr=" + encodeURIComponent(valueOf("sfinyr")) +
						"&sShSubID=" + encodeURIComponent(parts[0] || "") +
						"&sShSubSubID=" + encodeURIComponent(parts[1] || "") +
						"&id=0");
					childElements(root).forEach(function (node) {
						appendOption(level2, attrAt(node, 0), attrAt(node, 1));
					});
				}
			}
			return true;
		};

		window.setlevelfun1 = function () {
			var level3;
			var parts;
			var root;
			if (valueOf("sel2") === "A") {
				if (field("txtLev2")) {
					field("txtLev2").disabled = false;
				}
			} else {
				if (field("txtLev2")) {
					field("txtLev2").disabled = false;
				}
				setValue("txtLev2", selectedTextByName("sel2"));
				parts = splitCsv(valueOf("sel2"));
			}
			if (field("txtLev3")) {
				field("txtLev3").disabled = false;
			}
			level3 = resetSelect("sel3", "Add New", "A");
			if (field("sel3") && !field("sel3").disabled && valueOf("sel2") !== "A") {
				setValue("txtHierarchy", parts[1] || "");
				root = getXml("AccSubAndSubSubID.asp?sOrgID=" + encodeURIComponent(valueOf("sUnit")) +
					"&sFinyr=" + encodeURIComponent(valueOf("sfinyr")) +
					"&sBreakID=" + encodeURIComponent(parts[0] || "") +
					"&id=1");
				childElements(root).forEach(function (node) {
					appendOption(level3, attrAt(node, 0), [
						attrAt(node, 1),
						attrAt(node, 2),
						attrAt(node, 3),
						attrAt(node, 4),
						attrAt(node, 5),
						attrAt(node, 6),
						attrAt(node, 7),
						attrAt(node, 8)
					].join("-"));
				});
			}
			return true;
		};

		window.LevelFun = function () {
			var index = selectedRadioIndex("optLevel");
			if (index === 0) {
				setValue("txtLev2", "");
				setValue("txtAcHead", "");
				setValue("txtLev3", "");
				if (field("sel2")) {
					field("sel2").disabled = true;
				}
				if (field("sel3")) {
					field("sel3").disabled = true;
				}
				if (field("FinYear")) {
					field("FinYear").disabled = true;
				}
			} else if (index === 1) {
				setValue("txtLev2", "");
				setValue("txtLev3", "");
				if (field("sel2")) {
					field("sel2").disabled = false;
				}
				if (field("sel3")) {
					field("sel3").disabled = true;
				}
				if (field("FinYear")) {
					field("FinYear").disabled = true;
				}
			} else {
				setValue("txtLev2", "");
				setValue("txtLev3", "");
				if (field("sel2")) {
					field("sel2").disabled = false;
				}
				if (field("sel3")) {
					field("sel3").disabled = false;
				}
				if (field("FinYear")) {
					field("FinYear").disabled = true;
				}
			}
			return true;
		};

		window.ModeFun = function () {
			var accountMode = selectedRadioIndex("optMode") === 1;
			if (field("ButAcHead")) {
				field("ButAcHead").disabled = !accountMode;
			}
			if (field("FinYear")) {
				field("FinYear").disabled = !accountMode;
			}
			if (!accountMode) {
				setValue("txtAcHead", "");
			}
			return true;
		};

		window.AccHeadClck = function () {
			return true;
		};

		window.popAccList = function () {
			openPagedAccountHead("ChgAccHeadName.asp?orgId=" + encodeURIComponent(valueOf("sUnit")));
			return true;
		};

		window.Setlevelfun2 = function () {
			var parts;
			if (field("ButAcHead")) {
				field("ButAcHead").disabled = true;
			}
			setValue("hAccHead", "");
			setValue("txtAcHead", "");
			if (valueOf("sel3") === "A") {
				if (field("txtLev3")) {
					field("txtLev3").disabled = false;
				}
				return true;
			}
			if (field("txtLev3")) {
				field("txtLev3").disabled = false;
			}
			setValue("txtLev3", selectedTextByName("sel3"));
			parts = splitDash(valueOf("sel3"));
			if (parts[2] === "Y") {
				setRadioIndex("optMode", 1);
				if (field("ButAcHead")) {
					field("ButAcHead").disabled = false;
				}
				setValue("hAccHead", parts[4] || "");
				setValue("txtAcHead", parts[5] || "");
				if (field("FinYear")) {
					field("FinYear").disabled = false;
				}
			} else if (parts[3] !== "N") {
				setRadioIndex("optMode", 0);
			} else {
				setRadioIndex("optMode", 2);
			}
			if (parts[6] === "+" || parts[6] === "++") {
				setRadioIndex("optCompMode", parts[6] === "+" ? 0 : 1);
			}
			setValue("txtHierarchy", parts[7] || "");
			return true;
		};

		window.CheckSubmit = function () {
			var levelIndex = selectedRadioIndex("optLevel");
			var id = levelIndex === 0 ? "1" : levelIndex === 1 ? "2" : "3";
			var scheduleId = valueOf("sschedno");
			var scheduleSubId = "0";
			var scheduleSubSubId = "0";
			var breakupId = "0";
			var breakupSubId = "0";
			var breakupSubSubId = "";
			var headName = "";
			var breakupHeadName = "";
			var breakupSubHeadName = "";
			var parts;
			var root;
			var node;
			var response;
			if (id === "3" && valueOf("sel3") === "0") {
				alert("Select Level3 and Proceed..!");
				return false;
			}
			if (id === "2" && valueOf("sel2") !== "A") {
				headName = valueOf("txtLev1");
				breakupHeadName = valueOf("txtLev2");
				parts = splitDash(valueOf("sel1"));
				scheduleSubId = parts[0] || "0";
				scheduleSubSubId = parts[1] || "0";
			} else if (id === "3" && valueOf("sel3") !== "A") {
				breakupId = valueOf("sel2");
				parts = splitDash(valueOf("sel3"));
				breakupSubId = parts[0] || "0";
				breakupSubSubId = parts[1] || "0";
				breakupHeadName = valueOf("txtLev2");
				breakupSubHeadName = valueOf("txtLev3");
				parts = splitDash(valueOf("sel1"));
				scheduleSubId = parts[0] || "0";
				scheduleSubSubId = parts[1] || "0";
			} else {
				if (id === "1") {
					headName = valueOf("txtLev1");
				}
				if (id === "2") {
					if (valueOf("sel1") === "A") {
						headName = valueOf("txtLev1");
						breakupHeadName = valueOf("txtLev2");
					} else {
						breakupHeadName = valueOf("txtLev2");
						parts = splitDash(valueOf("sel1"));
						scheduleSubId = parts[0] || "0";
						scheduleSubSubId = parts[1] || "0";
					}
				}
				if (id === "3") {
					if (valueOf("sel2") === "A") {
						breakupHeadName = valueOf("txtLev2");
						breakupSubHeadName = valueOf("txtLev3");
						parts = splitDash(valueOf("sel1"));
						scheduleSubId = parts[0] || "0";
						scheduleSubSubId = parts[1] || "0";
					} else {
						breakupId = valueOf("sel2");
						breakupSubHeadName = valueOf("txtLev3");
					}
				}
			}
			root = tempRoot();
			node = createXmlElement("TempData", "Schedule");
			node.setAttribute("ID", id);
			node.setAttribute("Level2ID", valueOf("sel2"));
			node.setAttribute("Level3ID", valueOf("sel3"));
			node.setAttribute("OrgID", valueOf("sUnit"));
			node.setAttribute("ScheduleID", scheduleId);
			node.setAttribute("ScheduleSubID", scheduleSubId);
			node.setAttribute("ScheduleSubSubID", scheduleSubSubId);
			node.setAttribute("HeadName", headName);
			node.setAttribute("BreakUpHeadName", breakupHeadName);
			node.setAttribute("BreakUpSubHead", breakupSubHeadName);
			node.setAttribute("BreakupId", breakupId);
			node.setAttribute("BreakupSubId", breakupSubId);
			node.setAttribute("BreakupSubSubId", breakupSubSubId);
			node.setAttribute("Mode", selectedRadioIndex("optMode") === 1 ? "A" : "D");
			node.setAttribute("FinYear", valueOf("sfinyr"));
			node.setAttribute("ComputeMode", selectedRadioIndex("optCompMode") === 0 ? "+" : "++");
			node.setAttribute("AccountHeadID", valueOf("hAccHead"));
			node.setAttribute("Hierarchy", valueOf("txtHierarchy"));
			node.setAttribute("InsDate", valueOf("hInsDate"));
			root.appendChild(node);
			response = postText("XMLSchBrkHeadSave.asp?Name=SchedBrkSubHeads&Mod=Acc", serializeXml("TempData"));
			if (trim(response) !== "") {
				alert(response);
			} else {
				returnAndClose("Y");
			}
			return true;
		};

		window.Del = function () {
			var levelIndex = selectedRadioIndex("optLevel");
			var levelId = levelIndex === 1 ? "1" : levelIndex === 2 ? "2" : "";
			var breakupId = "0";
			var breakupSubId = "0";
			var breakupSubSubId = "0";
			var parts;
			var root;
			var node;
			var response;
			if (levelId === "1" && valueOf("sel2") !== "A") {
				breakupId = valueOf("sel2");
			} else if (levelId === "2" && valueOf("sel3") !== "A") {
				parts = splitDash(valueOf("sel3"));
				breakupId = valueOf("sel2");
				breakupSubId = parts[0] || "0";
				breakupSubSubId = parts[1] || "0";
			}
			root = tempRoot();
			node = createXmlElement("TempData", "Schedule");
			node.setAttribute("LevelID", levelId);
			node.setAttribute("iBreakID", breakupId);
			node.setAttribute("iBreakSubID", breakupSubId);
			node.setAttribute("iBreakSubSubID", breakupSubSubId);
			node.setAttribute("sOrgID", valueOf("sUnit"));
			node.setAttribute("sFinyr", valueOf("sfinyr"));
			node.setAttribute("sShID", valueOf("sschedno"));
			node.setAttribute("AcCode", valueOf("hAccHead"));
			node.setAttribute("id", "2");
			root.appendChild(node);
			response = postText("XMLShdDelete_Update.asp?Name=SchdDelete&Mod=Acc", serializeXml("TempData"));
			if (trim(response) !== "") {
				alert(response);
			} else {
				returnAndClose("Y");
			}
			return true;
		};

		window.populateSubID = function () {
			alert("1");
		};
	}

	function attrAt(node, index) {
		var item = node && node.attributes ? node.attributes.item(index) : null;
		return item ? item.nodeValue : "";
	}

	function addTableCell(row, className, align, content) {
		var cell = row.insertCell(-1);
		if (className) {
			cell.className = className;
		}
		if (align) {
			cell.align = align;
		}
		if (content != null) {
			if (typeof content === "string") {
				cell.innerHTML = content;
			} else {
				cell.appendChild(content);
			}
		}
		return cell;
	}

	function installTdsComputation() {
		function tdsTable() {
			return byId("tbltds");
		}

		function addTdsRow(serial, value, label, checked) {
			var row = tdsTable().insertRow(tdsTable().rows.length);
			var checkbox = document.createElement("input");
			checkbox.type = "checkbox";
			checkbox.name = "chkDel";
			checkbox.className = "ExcelDisplayCell";
			checkbox.value = value;
			checkbox.checked = !!checked;
			checkbox.onclick = window.CheckBoxClick;
			addTableCell(row, "ExcelHeaderCell", "center", String(serial));
			addTableCell(row, "ExcelDisplayCell", "center", checkbox);
			addTableCell(row, "ExcelDisplayCell", "left", label);
		}

		window.CheckBoxClick = function () {};
		window.Loadvalues = function () {
			var groupCode = valueOf("GroupCode");
			var headId = valueOf("HeadID");
			var xhr = new XMLHttpRequest();
			var root;
			var nodes;
			var formula = "";
			var formulaParts;
			var first;
			var voucher = "";
			var serial = toNumber(valueOf("iRowCount", "1"));
			xhr.open("GET", "TDSXMLGenerate.asp?GroupCode=" + encodeURIComponent(groupCode) + "&HeadID=" + encodeURIComponent(headId) + "&id=2", false);
			xhr.send(null);
			root = xhr.responseXML && xhr.responseXML.documentElement;
			if (!root && trim(xhr.responseText)) {
				root = new DOMParser().parseFromString(xhr.responseText, "text/xml").documentElement;
			}
			if (!root) {
				if (trim(xhr.responseText)) {
					alert(xhr.responseText);
				}
				return;
			}
			loadXmlIntoIsland("OutData", new XMLSerializer().serializeToString(root.ownerDocument || root));
			nodes = childElements(root);
			nodes.forEach(function (node) {
				formula = attrAt(node, 5);
			});
			formulaParts = formula ? formula.split(",") : [];
			if (formulaParts.length) {
				first = String(formulaParts[0]).split("-");
				voucher = first[0] || "";
				if (voucher === "0#0") {
					formulaParts.shift();
					first = String(formulaParts[0] || "").split("-");
				}
				if (first.length > 1) {
					setValue("txtpercentage", first[1]);
				}
			}
			addTdsRow(serial, "0#0", "Voucher Details", voucher === "0#0");
			serial += 1;
			nodes.forEach(function (node, index) {
				var merged = headId + "#" + attrAt(node, 0);
				var formulaItem = String(formulaParts[index] || "").split("-");
				var checkedToken = String((formulaItem[0] || "").split("#")[1] || "0");
				setValue("txtGname", attrAt(node, 6));
				setValue("txtcomputationfor", attrAt(node, 7));
				addTdsRow(serial, merged, attrAt(node, 1), checkedToken !== "0" && checkedToken !== "");
				serial += 1;
			});
			setValue("iRowCount", serial);
		};
		window.UpdateXML = window.Updatexml = function () {
			var groupCode = valueOf("GroupCode");
			var headId = valueOf("HeadID");
			var percentage = trim(valueOf("txtpercentage"));
			var checked = fields("chkDel").filter(function (checkbox) {
				return checkbox.checked;
			});
			var formulas = checked.map(function (checkbox) {
				return percentage ? checkbox.value + "-" + percentage : checkbox.value;
			});
			var doc = document.implementation.createDocument("", "Root", null);
			var node = doc.createElement("Schedule");
			var xhr = new XMLHttpRequest();
			node.setAttribute("GroupID", groupCode);
			node.setAttribute("HeadID", headId);
			node.setAttribute("ComputeFormula", formulas.join(","));
			doc.documentElement.appendChild(node);
			xhr.open("POST", "XMLcomputeSave.asp?Name=TDSComputeSave&Mod=Acc", false);
			xhr.send(new XMLSerializer().serializeToString(doc));
			if (trim(xhr.responseText) !== "") {
				alert(xhr.responseText);
			} else {
				returnAndClose("Y");
			}
		};
	}

	function installTdsGroupingSetup() {
		function tempRoot() {
			var root = xmlRoot("TempData");
			if (!root) {
				loadXmlIntoIsland("TempData", "<Root/>");
				root = xmlRoot("TempData");
			}
			clearChildren(root);
			return root;
		}

		function redirectToSetup() {
			var frm = form("formname");
			if (frm) {
				frm.action = "TDSGroupingSetup.asp?CallType=" + valueOf("hRequest");
				frm.submit();
			}
		}

		function rootFromDialogValue(value) {
			if (!value) {
				return null;
			}
			if (value.documentElement) {
				return value.documentElement;
			}
			if (value.XMLDocument && value.XMLDocument.documentElement) {
				return value.XMLDocument.documentElement;
			}
			if (value.nodeType === 1) {
				return value;
			}
			return null;
		}

		function getXml(url) {
			var xhr = new XMLHttpRequest();
			xhr.open("GET", url, false);
			xhr.send(null);
			if (xhr.responseXML && xhr.responseXML.documentElement) {
				loadXmlIntoIsland("OutData", new XMLSerializer().serializeToString(xhr.responseXML));
			} else if (trim(xhr.responseText) !== "") {
				loadXmlIntoIsland("OutData", xhr.responseText);
			}
			if (!xmlRoot("OutData") && trim(xhr.responseText)) {
				alert(xhr.responseText);
			}
			return xmlRoot("OutData");
		}

		function tdsSelectionRoot() {
			return xmlRoot("OutData");
		}

		function selectedTdsNode() {
			return firstElement(tdsSelectionRoot());
		}

		function selectedMode() {
			var radios = fields("R1");
			return radios[0] && radios[0].checked ? "F" : "P";
		}

		function setMode(value) {
			var radios = fields("R1");
			if (radios[0]) {
				radios[0].checked = value === "F";
			}
			if (radios[1]) {
				radios[1].checked = value !== "F";
			}
		}

		function postTemp(url) {
			return postText(url, serializeXml("TempData"));
		}

		function buildTdsUpdateDeleteNode(id, groupHeadId) {
			var root = tempRoot();
			var node = createXmlElement("TempData", "TDS");
			node.setAttribute("id", id);
			node.setAttribute("GroupID", valueOf("GroupName"));
			node.setAttribute("GroupHeadID", groupHeadId || "");
			node.setAttribute("GroupName", valueOf("TxtGroupName"));
			node.setAttribute("GroupHeadName", valueOf("TxtHead"));
			node.setAttribute("ComputeMode", selectedMode());
			node.setAttribute("AcHeadCode", valueOf("hAccHead"));
			node.setAttribute("Herarchy", valueOf("TxtHierachy"));
			node.setAttribute("AccHeadName", valueOf("TxtAccHeadName"));
			root.appendChild(node);
		}

		function continueGlSelection(url) {
			openModernDialog(url, xmlObject("GLHeadData"), "dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No", function (value) {
				var root = rootFromDialogValue(value);
				var action = trim(getAttribute(root, "Action")).toUpperCase();
				var query;
				var child;
				if (!root || action === "CLOSE") {
					return;
				}
				if (action && action !== "DONE") {
					query = trim(getAttribute(root, "PassQuery"));
					if (query) {
						continueGlSelection("../../Common/GLHeadSelection.asp?" + query);
					}
					return;
				}
				child = firstElement(root);
				if (child) {
					setValue("hAccHead", getAttribute(child, "RetField0"));
					setValue("TxtAccHeadName", getAttribute(child, "RetField5"));
				}
			});
		}

		window.AddGroup = function () {
			var callType = valueOf("hType");
			var root;
			var node;
			var response;
			if (callType === "E") {
				if (trim(valueOf("TxtGroupName")) === "" || trim(valueOf("GroupName")) === "") {
					alert("Group Name Required..!");
					if (field("TxtGroupName")) {
						field("TxtGroupName").focus();
					}
					return false;
				}
			} else if (callType === "C" && trim(valueOf("TxtGroupName")) === "") {
				alert("Group Name Required..!");
				if (field("TxtGroupName")) {
					field("TxtGroupName").focus();
				}
				return false;
			}
			if (trim(valueOf("TxtHead")) === "" || valueOf("selHead") === "0") {
				alert("Entet Head Name..!");
				if (field("TxtHead")) {
					field("TxtHead").focus();
				}
				return false;
			}
			if (trim(valueOf("TxtAccHeadName")) === "") {
				alert("Select Account Head..!");
				return false;
			}
			if (isNaN(Number(valueOf("TxtHierachy")))) {
				alert("Enter Numeric Value For Hierachy");
				setValue("TxtHierachy", "");
				return false;
			}
			setValue("hHeadName", valueOf("TxtGroupName"));
			root = tempRoot();
			node = createXmlElement("TempData", "Schedule");
			node.setAttribute("GroupID", callType === "E" ? valueOf("GroupName") : "A");
			node.setAttribute("GroupName", valueOf("TxtGroupName"));
			node.setAttribute("HeadID", valueOf("selHead"));
			node.setAttribute("HeadName", valueOf("TxtHead"));
			node.setAttribute("HeadCode", valueOf("hAccHead"));
			node.setAttribute("AccountHead", valueOf("TxtAccHeadName"));
			node.setAttribute("HeadDetails", valueOf("selHead"));
			node.setAttribute("Mode", selectedMode());
			node.setAttribute("Hierarchy", valueOf("TxtHierachy"));
			node.setAttribute("CreatedOn", "0");
			node.setAttribute("sOrgID", valueOf("OrgID"));
			root.appendChild(node);
			response = postTemp("XMLTDSSave.asp?Name=TDSDetails&Mod=Acc");
			if (trim(response) !== "") {
				alert(response);
			} else {
				redirectToSetup();
			}
			return true;
		};

		window.Submit = function () {
			var groupId = valueOf("GroupName");
			var countFields = fields("iTxtCount");
			var valueFields = fields("TxtVal");
			var formulaFields = fields("TxtFormula");
			var root;
			var response;
			var index;
			if (valueOf("TCount") === "0") {
				alert("Saved.. Sucessfully! ");
				return false;
			}
			root = tempRoot();
			for (index = 1; index < countFields.length; index += 1) {
				var node = createXmlElement("TempData", "TDS");
				node.setAttribute("id", "3");
				node.setAttribute("GroupID", groupId);
				node.setAttribute("TDSHeadID", valueFields[index] ? valueFields[index].value : "");
				node.setAttribute("Formula", formulaFields[index] ? formulaFields[index].value : "");
				root.appendChild(node);
			}
			response = postTemp("XMLTDSFormulaSave.asp?Name=TDSFormulaUpdate&Mod=Acc");
			if (trim(response) !== "") {
				alert(response);
			} else {
				alert("Saved...!");
				if (form("formname")) {
					form("formname").action = "TdsGroups.asp";
					form("formname").submit();
				}
			}
			return true;
		};

		window.SuppName = function () {
			continueGlSelection("../../Common/GLHeadSelection.asp?orgId=" + encodeURIComponent(valueOf("OrgID")));
			return true;
		};

		window.Calc = function (groupCode, headId) {
			openModernDialog("TDSComputationDetailPopup.asp?GroupCode=" + encodeURIComponent(groupCode) + "&HeadID=" + encodeURIComponent(headId), "A", "dialogHeight:320px;dialogWidth:710px;center:Yes;help:No;resizable:No;status:No", function () {});
			return true;
		};

		window.GNameChange = function () {
			if (valueOf("GroupName") !== "A") {
				redirectToSetup();
			} else {
				setValue("TxtGroupName", "");
			}
			return true;
		};

		window.UnitChange = function (type) {
			var select = field("selGPName");
			var root;
			if (type !== "E" || !select) {
				return true;
			}
			select.options.length = 0;
			root = getXml("TDSXMLGenerate.asp?sOrgID=" + encodeURIComponent(valueOf("OrgID")) + "&id=0");
			childElements(root).forEach(function (node) {
				select.options[select.options.length] = new Option(attrAt(node, 0), attrAt(node, 1));
			});
			return true;
		};

		window.Del = function () {
			return true;
		};

		window.GChange = function () {
			return true;
		};

		window.ShowVouch = function (headId) {
			var root;
			var node;
			if (valueOf("hType") === "C") {
				return false;
			}
			root = getXml("TDSXMLGenerate.asp?HeadID=" + encodeURIComponent(headId) + "&GroupID=" + encodeURIComponent(valueOf("GroupName")) + "&id=3");
			node = firstElement(root);
			if (!node) {
				return false;
			}
			if (field("selHead")) {
				field("selHead").selectedIndex = 0;
				field("selHead").disabled = true;
			}
			setValue("TxtGroupName", attrAt(node, 6));
			setValue("TxtHead", attrAt(node, 7));
			setValue("TxtAccHeadName", attrAt(node, 8));
			setValue("hAccHead", attrAt(node, 3));
			setValue("TxtHierachy", attrAt(node, 4));
			setMode(attrAt(node, 2));
			if (field("ButUpdate")) {
				field("ButUpdate").disabled = false;
			}
			if (field("ButDelete")) {
				field("ButDelete").disabled = false;
			}
			if (field("ButAdd")) {
				field("ButAdd").disabled = true;
			}
			return true;
		};

		window.UpdateGroup = function () {
			var node = selectedTdsNode();
			var response;
			if (!node) {
				alert("Select TDS head to update.");
				return false;
			}
			buildTdsUpdateDeleteNode("1", attrAt(node, 0));
			response = postTemp("XMLTDSUpdateDelete.asp?Name=TDSUpdateDelete&Mod=Acc");
			if (trim(response) !== "") {
				alert(response);
			} else {
				redirectToSetup();
			}
			return true;
		};

		window.DeleteGroup = function () {
			var node = selectedTdsNode();
			var response;
			if (!node) {
				alert("Select TDS head to delete.");
				return false;
			}
			buildTdsUpdateDeleteNode("2", attrAt(node, 0));
			response = postTemp("XMLTDSUpdateDelete.asp?Name=TDSUpdateDelete&Mod=Acc");
			if (trim(response) !== "") {
				alert(response);
			} else {
				setValue("TxtGroupName", "");
				setValue("TxtHead", "");
				redirectToSetup();
			}
			return true;
		};

		window.TDSDel = function () {
			var frm = form("formname");
			if (frm) {
				frm.action = "TDSGroupingDelete.asp";
				frm.submit();
			}
			return true;
		};
	}

	function installBankBookDetailsPopup() {
		function setText(id, value) {
			var item = byId(id);
			if (item) {
				item.textContent = value == null ? "" : String(value);
			}
		}

		function setSelectValue(name, value, fallbackIndex) {
			var select = field(name);
			if (!select) {
				return;
			}
			select.value = value;
			if (select.value !== String(value) && fallbackIndex != null) {
				select.selectedIndex = fallbackIndex;
			}
		}

		function checkedRadioValue(name, fallback) {
			var checked = fields(name).filter(function (radio) { return radio.checked; })[0];
			return checked ? checked.value : fallback || "";
		}

		function tempRoot() {
			var root = xmlRoot("BankBookDet");
			if (!root) {
				loadXmlIntoIsland("BankBookDet", "<Root/>");
				root = xmlRoot("BankBookDet");
			}
			clearChildren(root);
			return root;
		}

		function dialogRootFromValue(value) {
			if (!value) {
				return null;
			}
			if (value.documentElement) {
				return value.documentElement;
			}
			if (value.XMLDocument && value.XMLDocument.documentElement) {
				return value.XMLDocument.documentElement;
			}
			if (value.nodeType === 1) {
				return value;
			}
			return null;
		}

		function continueGlSelection(url, flag) {
			openModernDialog(url, xmlObject("GLHeadData"), "dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No", function (value) {
				var root = dialogRootFromValue(value);
				var action = trim(getAttribute(root, "Action")).toUpperCase();
				var query;
				var child;
				if (!root || action === "CLOSE") {
					return;
				}
				if (action && action !== "DONE") {
					query = trim(getAttribute(root, "PassQuery"));
					if (query) {
						continueGlSelection("../../Common/GLHeadSelection.asp?" + query, flag);
					}
					return;
				}
				child = firstElement(root);
				if (!child) {
					return;
				}
				if (flag === "C") {
					setValue("hChargestHead", getAttribute(child, "RetField0"));
					setText("spCharges", getAttribute(child, "RetField5"));
				} else {
					setValue("hDiscountHead", getAttribute(child, "RetField0"));
					setText("spDisCount", getAttribute(child, "RetField5"));
				}
			});
		}

		window.popBankBook = function () {
			var select = field("selBankBook");
			var root = xmlRoot("BookData");
			if (!select || !root) {
				return false;
			}
			childElements(root).forEach(function (node) {
				select.options[select.options.length] = new Option(attrAt(node, 1), attrAt(node, 0));
			});
			return true;
		};

		window.DisplayBookDet = function () {
			var xhr = new XMLHttpRequest();
			var text;
			var root;
			xhr.open("GET", "XMLBankBookDetail.asp?orgID=" + encodeURIComponent(valueOf("hOrgID")) + "&BookNo=" + encodeURIComponent(valueOf("hBookNo")), false);
			xhr.send(null);
			text = xhr.responseXML && xhr.responseXML.documentElement ? new XMLSerializer().serializeToString(xhr.responseXML) : xhr.responseText;
			root = trim(text) ? loadXmlIntoIsland("BookData", text) : null;
			if (root && childElements(root).length) {
				setValue("hActionFlag", "U");
				window.popBankBookDetail();
			} else {
				if (form("formname")) {
					form("formname").reset();
				}
				setValue("hActionFlag", "I");
				setText("spCharges", "");
				setText("spDisCount", "");
			}
			return true;
		};

		window.popBankBookDetail = function () {
			var node = firstElement(xmlRoot("BookData"));
			var cheque = fields("optCheque");
			var payIn = fields("optPayIn");
			if (!node) {
				return false;
			}
			setValue("txtName", getAttribute(node, "BankName"));
			setValue("txtAddress1", getAttribute(node, "BankAddress1"));
			setValue("txtAddress2", getAttribute(node, "BankAddress2"));
			setValue("txtCity", getAttribute(node, "City"));
			setValue("txtState", getAttribute(node, "State"));
			setValue("txtCountry", getAttribute(node, "Country"));
			setValue("txtPinCode", getAttribute(node, "PinCode"));
			setValue("txtPhone", getAttribute(node, "PhoneNos"));
			setValue("txtMobileNo", getAttribute(node, "MobileNos"));
			setValue("txtFax", getAttribute(node, "FaxNos"));
			setValue("txtEmail", getAttribute(node, "EMailId"));
			setValue("txtWebsite", getAttribute(node, "WebSiteURL"));
			if (cheque[0]) {
				cheque[0].checked = getAttribute(node, "PrintCheques") === "1";
			}
			if (cheque[1]) {
				cheque[1].checked = getAttribute(node, "PrintCheques") !== "1";
			}
			if (payIn[0]) {
				payIn[0].checked = getAttribute(node, "PrintPayInSlip") === "1";
			}
			if (payIn[1]) {
				payIn[1].checked = getAttribute(node, "PrintPayInSlip") !== "1";
			}
			if (getAttribute(node, "AccountType") === "CU") {
				setValue("txtCreditLimit", "0");
				if (field("txtCreditLimit")) {
					field("txtCreditLimit").readOnly = true;
				}
				setSelectValue("selAccType", "CU", 1);
			} else {
				if (field("txtCreditLimit")) {
					field("txtCreditLimit").readOnly = false;
				}
				setSelectValue("selAccType", "CC", 2);
			}
			setValue("txtAccNo", getAttribute(node, "AccountNo"));
			setValue("txtCreditLimit", getAttribute(node, "CreditLimit"));
			setValue("txtODLimit", getAttribute(node, "OverDraftLimit"));
			setValue("txtDiscountLimit", getAttribute(node, "DiscountingLimit"));
			setValue("txtLCLimit", getAttribute(node, "LCLimit"));
			setValue("txtswitCode", getAttribute(node, "SwiftCode"));
			setValue("hChargestHead", getAttribute(node, "ChargeHead"));
			setText("spCharges", getAttribute(node, "ChargeHeadName"));
			setValue("hDiscountHead", getAttribute(node, "DiscountHead"));
			setText("spDisCount", getAttribute(node, "DiscountHeadName"));
			return true;
		};

		window.popAccList = function (flag) {
			continueGlSelection("../../Common/GLHeadSelection.asp?orgId=" + encodeURIComponent(valueOf("hOrgID")), flag);
			return true;
		};

		window.CheckSubmit = function () {
			var numericFields = ["txtCreditLimit", "txtODLimit", "txtLCLimit", "txtDiscountLimit"];
			var root;
			var node;
			var response;
			var index;
			if (field("selAccType") && field("selAccType").selectedIndex < 1) {
				alert("Select Account Type");
				field("selAccType").focus();
				return false;
			}
			if (trim(valueOf("txtName")) === "") {
				alert("Enter Bank Name");
				field("txtName").focus();
				return false;
			}
			if (trim(valueOf("txtAccNo")) === "") {
				alert("Enter Account No");
				field("txtAccNo").focus();
				return false;
			}
			for (index = 0; index < numericFields.length; index += 1) {
				var item = field(numericFields[index]);
				var value = toNumber(valueOf(numericFields[index]));
				if (isNaN(parseFloat(valueOf(numericFields[index])))) {
					alert("Enter Numeric Value ");
					setValue(numericFields[index], "0");
					if (item && item.select) {
						item.select();
					}
					return false;
				}
				if (value < 0) {
					alert("Enter Value Greater than or Equal to Zero");
					setValue(numericFields[index], "0");
					if (item && item.select) {
						item.select();
					}
					return false;
				}
			}
			root = tempRoot();
			node = createXmlElement("BankBookDet", "BankBook");
			node.setAttribute("UnitID", valueOf("hOrgID"));
			node.setAttribute("BookNo", valueOf("hBookNo"));
			node.setAttribute("ActionFlag", valueOf("hActionFlag"));
			node.setAttribute("BankName", valueOf("txtName"));
			node.setAttribute("Address1", valueOf("txtAddress1"));
			node.setAttribute("Address2", valueOf("txtAddress2"));
			node.setAttribute("City", valueOf("txtCity"));
			node.setAttribute("State", valueOf("txtState"));
			node.setAttribute("Country", valueOf("txtCountry"));
			node.setAttribute("Pincode", valueOf("txtPinCode"));
			node.setAttribute("Phone", valueOf("txtPhone"));
			node.setAttribute("MobileNo", valueOf("txtMobileNo"));
			node.setAttribute("Fax", valueOf("txtFax"));
			node.setAttribute("EMail", valueOf("txtEmail"));
			node.setAttribute("WebSite", valueOf("txtWebsite"));
			node.setAttribute("PrintCheque", checkedRadioValue("optCheque", "0"));
			node.setAttribute("PrintPayInSlip", checkedRadioValue("optPayIn", "0"));
			node.setAttribute("AccountType", valueOf("selAccType"));
			node.setAttribute("AccountNo", valueOf("txtAccNo"));
			node.setAttribute("CreditLimit", valueOf("txtCreditLimit"));
			node.setAttribute("ODLimit", valueOf("txtODLimit"));
			node.setAttribute("DiscountLimit", valueOf("txtDiscountLimit"));
			node.setAttribute("LCLimit", valueOf("txtLCLimit"));
			node.setAttribute("SwiftCode", valueOf("txtswitCode"));
			node.setAttribute("ChargesHead", valueOf("hChargestHead"));
			node.setAttribute("DiscountHead", valueOf("hDiscountHead"));
			root.appendChild(node);
			response = postText("BankBookDetailsUpdate.asp", serializeXml("BankBookDet"));
			if (trim(response) !== "") {
				alert(response);
				return false;
			}
			alert("Bank Details Inserted Successfully");
			returnAndClose("Done");
			return true;
		};

		window.checkCredit = function () {
			if (field("selAccType") && field("selAccType").selectedIndex === 2) {
				if (field("txtCreditLimit")) {
					field("txtCreditLimit").readOnly = false;
				}
			} else {
				setValue("txtCreditLimit", "0");
				if (field("txtCreditLimit")) {
					field("txtCreditLimit").readOnly = true;
				}
			}
			return true;
		};
	}

	function installDayBookGrid() {
		function submitForm() {
			if (form("formname")) {
				form("formname").submit();
			}
		}

		function dialogRootFromValue(value) {
			if (!value) {
				return null;
			}
			if (value.documentElement) {
				return value.documentElement;
			}
			if (value.XMLDocument && value.XMLDocument.documentElement) {
				return value.XMLDocument.documentElement;
			}
			if (value.nodeType === 1) {
				return value;
			}
			return null;
		}

		function selectGlHead(orgId, callback) {
			function open(url) {
				openModernDialog(url, xmlObject("GLHeadData"), "dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No", function (value) {
					var root = dialogRootFromValue(value);
					var action = trim(getAttribute(root, "Action")).toUpperCase();
					var query;
					var child;
					if (!root || action === "CLOSE") {
						callback("");
						return;
					}
					if (action && action !== "DONE") {
						query = trim(getAttribute(root, "PassQuery"));
						if (query) {
							open("../../Common/GLHeadSelection.asp?" + query);
						}
						return;
					}
					child = firstElement(root);
					callback(child ? getAttribute(child, "RetField0") : "");
				});
			}
			open("../../Common/GLHeadSelection.asp?orgId=" + encodeURIComponent(orgId));
		}

		function updateGlHead(orgId, bookNo, bookId, glHead) {
			var xhr = new XMLHttpRequest();
			var text;
			if (!glHead || toNumber(glHead) <= 0) {
				return;
			}
			xhr.open("GET", "GetAccountHeadMaped.asp?OrgCode=" + encodeURIComponent(orgId) + "&AccHead=" + encodeURIComponent(glHead), false);
			xhr.send(null);
			if (trim(xhr.responseText) !== "" && toNumber(xhr.responseText) > 0) {
				alert("Selected account head has already been mapped to another daybook. Please select some other account head");
				selectGlHead(orgId, function (nextHead) {
					updateGlHead(orgId, bookNo, bookId, nextHead);
				});
				return;
			}
			xhr = new XMLHttpRequest();
			xhr.open("GET", "BooksCodeUpdate.asp?orgID=" + encodeURIComponent(orgId) + "&BookId=" + encodeURIComponent(bookId) + "&BookNo=" + encodeURIComponent(bookNo) + "&GlHead=" + encodeURIComponent(glHead), false);
			xhr.send(null);
			text = xhr.responseXML && xhr.responseXML.documentElement ? new XMLSerializer().serializeToString(xhr.responseXML) : xhr.responseText;
			if (trim(text) !== "") {
				loadXmlIntoIsland("OutData", text);
				submitForm();
			}
		}

		window.CreateNewParty = function () {
			if (form("formname")) {
				form("formname").action = "BooksCreationEntry.asp";
				form("formname").submit();
			}
			return true;
		};
		window.AssignPage = function (page) {
			setValue("hPage", page);
			submitForm();
			return true;
		};
		window.CheckSubmit = function () {
			setValue("hBookName", valueOf("txtBookName"));
			setValue("hBookType", valueOf("selBookType"));
			setValue("hBookTypeName", selectedText(field("selBookType")));
			submitForm();
			return true;
		};
		window.EditBook = function (orgCode, bookCode, bookNumber, fromAccHead) {
			openModernDialog("BooksEditEntryPopup.asp?OrgCode=" + encodeURIComponent(orgCode) + "&BookCode=" + encodeURIComponent(bookCode) + "&BookNumber=" + encodeURIComponent(bookNumber) + "&FromAcc=" + encodeURIComponent(fromAccHead), "", "Status:No;", function (value) {
				if (value === "Done") {
					submitForm();
				}
			});
			return true;
		};
		window.ShowGLHead = function (orgId) {
			selectGlHead(orgId, function () {});
			return "";
		};
		window.GlChange = function (orgId, bookNo, bookId) {
			selectGlHead(orgId, function (glHead) {
				updateGlHead(orgId, bookNo, bookId, glHead);
			});
			return true;
		};
		window.ViewContraDet = function (orgCode, fromAccHead) {
			openModernDialog("ContraListPopup.asp?OrgCode=" + encodeURIComponent(orgCode) + "&FromAcc=" + encodeURIComponent(fromAccHead), "", "dialogHeight:430px;dialogWidth:465px;Status:No;Help:No;", function (value) {
				if (value === "Done") {
					submitForm();
				}
			});
			return true;
		};
		window.ShowBankBookDet = function (orgCode, bookCode, bookNumber, fromAccHead) {
			openModernDialog("BankBookDetailsPopup.asp?OrgCode=" + encodeURIComponent(orgCode) + "&BookCode=" + encodeURIComponent(bookCode) + "&BookNumber=" + encodeURIComponent(bookNumber) + "&FromAcc=" + encodeURIComponent(fromAccHead), "", "Status:No;", function () {});
			return true;
		};
	}

	function installBookNarrations() {
		function submitForm() {
			if (form("formname")) {
				form("formname").submit();
			}
		}

		function dialogRootFromValue(value) {
			if (!value) {
				return null;
			}
			if (value.documentElement) {
				return value.documentElement;
			}
			if (value.XMLDocument && value.XMLDocument.documentElement) {
				return value.XMLDocument.documentElement;
			}
			if (value.nodeType === 1) {
				return value;
			}
			return null;
		}

		function openNarrationPopup(type, narrNo) {
			var url = "NarrationEntryPopUp.asp?BookCode=" + encodeURIComponent(valueOf("hBookCode")) + "&Type=" + encodeURIComponent(type);
			if (narrNo != null) {
				url += "&NarrNo=" + encodeURIComponent(narrNo);
			}
			openModernDialog(url, "", "dialogHeight:190px;Status:no", function (value) {
				var root = dialogRootFromValue(value);
				if (getAttribute(root, "Done") === "Y") {
					submitForm();
				}
			});
		}

		window.ViewContactDeatils = function (partyCode) {
			openModernDialog("ParDisplayContactDetails.asp?PartyCode=" + encodeURIComponent(partyCode), "", "dialogHeight:350px;Status:no", function () {});
			return true;
		};
		window.CreateNew = function () {
			openNarrationPopup("N");
			return true;
		};
		window.DeleteData = function () {
			var count = toNumber(valueOf("hCnt"));
			var selected = [];
			var index;
			for (index = 1; index <= count; index += 1) {
				var checkbox = field("chkbox" + index);
				var parts;
				if (checkbox && checkbox.checked && checkbox.value) {
					parts = String(checkbox.value).split(":");
					selected.push(parts[0]);
				}
			}
			if (!selected.length) {
				alert("Select any one Narration For Delete");
				return false;
			}
			if (form("formname")) {
				form("formname").action = "NarrationDelete.asp?NarrationNo=" + encodeURIComponent(selected.join(",")) + "&BookCode=" + encodeURIComponent(valueOf("hBookCode"));
				form("formname").submit();
			}
			return true;
		};
		window.AssignPage = function (page) {
			setValue("hPage", page);
			submitForm();
			return true;
		};
		window.CheckSubmit = function () {
			setValue("hDayBook", valueOf("selDayBook"));
			setValue("hNarration", valueOf("txtNarration"));
			submitForm();
			return true;
		};
		window.EditNarration = function (narrNo) {
			openNarrationPopup("E", narrNo);
			return true;
		};
	}

	function installNarrationEntryPopup() {
		function ensureRoot(name) {
			var root = xmlRoot(name);
			if (!root) {
				loadXmlIntoIsland(name, name === "RetData" ? "<ROOT Done=\"\"/>" : "<Root/>");
				root = xmlRoot(name);
			}
			return root;
		}

		function dialogRootFromValue(value) {
			if (!value) {
				return null;
			}
			if (value.documentElement) {
				return value.documentElement;
			}
			if (value.XMLDocument && value.XMLDocument.documentElement) {
				return value.XMLDocument.documentElement;
			}
			if (value.nodeType === 1) {
				return value;
			}
			return null;
		}

		function continueBookSelection(url) {
			openModernDialog(url, xmlObject("BookDet"), "dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No", function (value) {
				var root = dialogRootFromValue(value);
				var action = trim(getAttribute(root, "Action")).toUpperCase();
				var query;
				var bookCodes = [];
				var bookNos = [];
				var bookNames = [];
				if (!root || action === "CLOSE") {
					return;
				}
				if (action && action !== "DONE") {
					query = trim(getAttribute(root, "PassQuery"));
					continueBookSelection(query ? "BookSelectionPopUp.asp?" + query : url);
					return;
				}
				childElements(root, "Entry").forEach(function (node) {
					bookCodes.push(getAttribute(node, "RetField0"));
					bookNos.push(getAttribute(node, "RetField1"));
					bookNames.push(getAttribute(node, "RetField2"));
				});
				setValue("hSelBookCode", bookCodes.join(","));
				setValue("hSelBookNo", bookNos.join(","));
				if (byId("UsedInBook")) {
					byId("UsedInBook").textContent = bookNames.join(",");
				}
			});
		}

		window.checkSubmit = window.CheckSubmit = function () {
			var outRoot = ensureRoot("OutData");
			var retRoot = ensureRoot("RetData");
			var node;
			var response;
			if (trim(valueOf("txtShortDesc")) === "") {
				alert("Enter Narration Short Description");
				if (field("txtShortDesc")) {
					field("txtShortDesc").focus();
				}
				return false;
			}
			if (trim(valueOf("txtDesc")) === "") {
				alert("Enter Narration Description");
				if (field("txtDesc")) {
					field("txtDesc").focus();
				}
				return false;
			}
			clearChildren(outRoot);
			node = createXmlElement("OutData", "Desc");
			node.setAttribute("Type", valueOf("hType"));
			node.setAttribute("ShortDesc", valueOf("txtShortDesc"));
			node.setAttribute("Desc", valueOf("txtDesc"));
			node.setAttribute("BookCode", valueOf("hSelBookCode"));
			node.setAttribute("BookNo", valueOf("hSelBookNo"));
			node.setAttribute("NarrNo", valueOf("hNarrNo"));
			outRoot.appendChild(node);
			response = postText("NarrationUpdate.asp", serializeXml("OutData"));
			if (trim(response) !== "") {
				alert(response);
				return false;
			}
			alert(valueOf("hType") === "N" ? "Narration Details Inserted Successfully" : "Narration Details Updated Successfully");
			retRoot.setAttribute("Done", "Y");
			returnAndClose(retRoot);
			return true;
		};

		window.selBook = function () {
			continueBookSelection("BookSelectionPopUp.asp?hSelectMode=M&BookCode=" + encodeURIComponent(valueOf("hBookCode")));
			return true;
		};
	}

	function installBooksEditEntryPopup() {
		var submitting = false;

		function setImage(id, src) {
			var image = byId(id);
			if (image) {
				image.src = src;
			}
		}

		function radios(name) {
			return fields(name);
		}

		function selectedRadioValue(name, fallback) {
			var selected = radios(name).filter(function (radio) { return radio.checked; })[0];
			return selected ? selected.value : fallback || "";
		}

		function selectedSeriesNode() {
			var root = xmlRoot("SeriesNoData");
			var seriesNo = valueOf("selNoSeries");
			return childElements(root, "Series").filter(function (node) {
				return attrAt(node, 0) === seriesNo;
			})[0] || null;
		}

		function periodText(seriesNode, entryNode) {
			var period = attrAt(entryNode, 1);
			var counterType = attrAt(seriesNode, 3);
			if (counterType === "M") {
				return "Month-" + period;
			}
			if (counterType === "Q") {
				return "Quater-" + period;
			}
			if (counterType === "Y") {
				return "Yearly";
			}
			return period;
		}

		function addTextCell(row, name, value, size, maxLength) {
			var input = document.createElement("input");
			input.type = "text";
			input.name = name || "";
			input.value = value == null ? "" : String(value);
			input.className = "Formelem";
			if (size) {
				input.size = size;
			}
			if (maxLength) {
				input.maxLength = maxLength;
			}
			addTableCell(row, "ExcelInputCell", "", input);
		}

		function clearBookTable(flag) {
			var table = byId("tblBook");
			var row;
			if (!table) {
				return;
			}
			while (table.rows.length) {
				table.deleteRow(0);
			}
			row = table.insertRow(0);
			addTableCell(row, "ExcelSerial", "Center", "S.No");
			addTableCell(row, "ExcelHeaderCell", "left", "Period");
			if (flag === "Y") {
				addTableCell(row, "ExcelHeaderCell", "left", "CR StartNo");
				addTableCell(row, "ExcelHeaderCell", "left", "CR Prefix");
				addTableCell(row, "ExcelHeaderCell", "left", "CR Suffix");
				addTableCell(row, "ExcelHeaderCell", "left", "DR StartNo");
				addTableCell(row, "ExcelHeaderCell", "left", "DR Prefix");
				addTableCell(row, "ExcelHeaderCell", "left", "DR Suffix");
			} else {
				addTableCell(row, "ExcelHeaderCell", "left", "StartNo");
				addTableCell(row, "ExcelHeaderCell", "left", "Prefix");
				addTableCell(row, "ExcelHeaderCell", "left", "Suffix");
			}
		}

		function renderContraBookRows(root) {
			var table = byId("tblMap");
			var index = 1;
			if (!table) {
				return;
			}
			while (table.rows.length > 1) {
				table.deleteRow(1);
			}
			childElements(root).forEach(function (node) {
				var row = table.insertRow(table.rows.length);
				var checkbox = document.createElement("input");
				checkbox.type = "checkbox";
				checkbox.name = "chkBox" + index;
				checkbox.className = "FormElem";
				checkbox.value = getAttribute(node, "No");
				checkbox.disabled = getAttribute(node, "Records") === "Y";
				addTableCell(row, "ExcelSerial", "center", String(index));
				addTableCell(row, "ExcelDisplayCell", "center", checkbox);
				addTableCell(row, "ExcelDisplayCell", "left", getAttribute(node, "Name"));
				index += 1;
			});
			setValue("hRowContraCnt", index - 1);
		}

		function refreshContraBookRows() {
			var xhr = new XMLHttpRequest();
			var root;
			xhr.open("GET", "GetMappedContraDetails.asp?OrgCode=" + encodeURIComponent(valueOf("hOrgCode")) + "&FromHead=" + encodeURIComponent(valueOf("hFromHead")), false);
			xhr.send(null);
			root = xhr.responseXML && xhr.responseXML.documentElement;
			if (!root && trim(xhr.responseText)) {
				root = new DOMParser().parseFromString(xhr.responseText, "text/xml").documentElement;
			}
			if (!root) {
				if (trim(xhr.responseText)) {
					alert(xhr.responseText);
				}
				return;
			}
			loadXmlIntoIsland("AccData", new XMLSerializer().serializeToString(root.ownerDocument || root));
			renderContraBookRows(root);
		}

		window.DivClick = function (value) {
			var divContra = byId("DivContra");
			var divBasic = byId("DivBasic");
			if (value === "DivContra" && divContra) {
				if (divContra.style.display === "block") {
					divContra.style.display = "none";
					setImage("imgContra", "../../assets/images/plus.gif");
				} else {
					divContra.style.display = "block";
					setImage("imgContra", "../../assets/images/minus.gif");
					if (divBasic) {
						divBasic.style.display = "none";
					}
					setImage("imgBasic", "../../assets/images/plus.gif");
				}
			} else if (value === "DivBasic" && divBasic) {
				if (divBasic.style.display === "block") {
					divBasic.style.display = "none";
					setImage("imgBasic", "../../assets/images/plus.gif");
				} else {
					divBasic.style.display = "block";
					setImage("imgBasic", "../../assets/images/minus.gif");
					if (trim(valueOf("hFromHead")) !== "0" && divContra) {
						divContra.style.display = "none";
						setImage("imgContra", "../../assets/images/plus.gif");
					}
				}
			}
			return true;
		};

		window.FunUseable = function () {
			var xhr = new XMLHttpRequest();
			var useable;
			xhr.open("GET", "GetBookUsedDetails.asp?OrgCode=" + encodeURIComponent(valueOf("hOrgCode")) + "&BookCode=" + encodeURIComponent(valueOf("hBookCode")) + "&BookNo=" + encodeURIComponent(valueOf("hBookNo")), false);
			xhr.send(null);
			if (trim(xhr.responseText) === "Y") {
				alert("The Current Book is Already having transactions could not be change the Useable Status");
				if (form("formname")) {
					form("formname").submit();
				}
				return false;
			}
			useable = radios("optUseable")[0] && radios("optUseable")[0].checked ? "0" : "1";
			if (!confirm("Do you want to change the Book Useable?")) {
				return false;
			}
			xhr = new XMLHttpRequest();
			xhr.open("GET", "BookUseableUpdate.asp?OrgCode=" + encodeURIComponent(valueOf("hOrgCode")) + "&BookCode=" + encodeURIComponent(valueOf("hBookCode")) + "&BookNo=" + encodeURIComponent(valueOf("hBookNo")) + "&Useable=" + encodeURIComponent(useable), false);
			xhr.send(null);
			if (trim(xhr.responseText) !== "") {
				if (xhr.responseText.length > 1) {
					alert(xhr.responseText);
					return false;
				}
				useable = toNumber(xhr.responseText);
				if (radios("optUseable")[0]) {
					radios("optUseable")[0].checked = useable === 0;
				}
				if (radios("optUseable")[1]) {
					radios("optUseable")[1].checked = useable !== 0;
				}
			}
			return true;
		};

		window.FormClose = function () {
			setValue("hAction", "Close");
			returnAndClose("Close");
			return true;
		};

		window.DisplayBook = function () {
			var seriesNode = selectedSeriesNode();
			var entries;
			var index = 1;
			var separate = valueOf("selPayRecNo");
			clearBookTable(separate);
			if (!seriesNode || (field("selNoSeries") && field("selNoSeries").selectedIndex === 0)) {
				setValue("hRowCnt", "1");
				return false;
			}
			entries = childElements(seriesNode, "Entry");
			if (entries.length === 12) {
				setValue("hSeriesType", "M");
			} else if (entries.length === 4) {
				setValue("hSeriesType", "Q");
			} else if (entries.length === 1) {
				setValue("hSeriesType", "Y");
			} else {
				setValue("hSeriesType", attrAt(seriesNode, 2));
			}
			setValue("hSeriesLen", attrAt(seriesNode, 4));
			entries.forEach(function (entry) {
				var row = byId("tblBook").insertRow(byId("tblBook").rows.length);
				var entryNo = attrAt(entry, 0);
				addTableCell(row, "ExcelSerial", "Center", String(index));
				addTableCell(row, "ExcelDisplayCell", "left", periodText(seriesNode, entry));
				if (separate === "Y") {
					addTextCell(row, "txtCrStartNo" + entryNo, "1", 5, 7);
					addTextCell(row, "txtCrPrefix" + entryNo, attrAt(entry, 3), 11, 10);
					addTextCell(row, "txtCrSuffix" + entryNo, attrAt(entry, 4), 11, 10);
					addTextCell(row, "txtDrStartNo" + entryNo, "1", 5, 7);
					addTextCell(row, "txtDrPrefix" + entryNo, attrAt(entry, 3), 11, 10);
					addTextCell(row, "txtDrSuffix" + entryNo, attrAt(entry, 4), 11, 10);
				} else {
					addTextCell(row, "txtStartNo" + entryNo, "1", 5, 7);
					addTextCell(row, "txtPrefix" + entryNo, attrAt(entry, 3), 11, 10);
					addTextCell(row, "txtSuffix" + entryNo, attrAt(entry, 4), 11, 10);
				}
				index += 1;
			});
			setValue("hRowCnt", index);
			return true;
		};

		window.popSeriesNo = function () {
			var root = xmlRoot("SeriesNoData");
			var select = field("selNoSeries");
			var wanted = valueOf("hCounterType");
			if (!root || !select) {
				return false;
			}
			select.options.length = 0;
			select.options[0] = new Option("Select Number Series", "0");
			childElements(root).forEach(function (node) {
				if (attrAt(node, 3) === "M") {
					select.options[select.options.length] = new Option(attrAt(node, 1), attrAt(node, 0));
				}
			});
			Array.prototype.forEach.call(select.options, function (option, index) {
				if (String(option.value) === String(wanted)) {
					select.selectedIndex = index;
				}
			});
			return true;
		};

		window.ClearTable1 = clearBookTable;

		window.validateForm = function () {
			var seriesNode;
			var entries;
			if (trim(valueOf("txtName")) === "") {
				alert("Enter Book Name");
				if (field("txtName")) {
					field("txtName").focus();
				}
				return false;
			}
			if (field("selNoSeries") && field("selNoSeries").selectedIndex === 0) {
				alert("Select No Series ");
				field("selNoSeries").focus();
				return false;
			}
			if (trim(valueOf("hSeriesType")) === "") {
				seriesNode = selectedSeriesNode();
				entries = childElements(seriesNode, "Entry");
				if (seriesNode) {
					setValue("hSeriesType", entries.length === 12 ? "M" : entries.length === 4 ? "Q" : entries.length === 1 ? "Y" : attrAt(seriesNode, 2));
					setValue("hSeriesLen", attrAt(seriesNode, 4));
				}
			}
			setValue("hAction", "Done");
			setValue("hCallType", "E");
			if (field("B3")) {
				field("B3").disabled = true;
			}
			if (field("B4")) {
				field("B4").disabled = true;
			}
			if (form("formname")) {
				submitting = true;
				form("formname").action = "BooksEditPopupUpdate.asp";
				form("formname").submit();
			}
			return true;
		};

		window.DelBook = function () {
			setValue("hAction", "Done");
			setValue("hCallType", "D");
			if (field("B3")) {
				field("B3").disabled = true;
			}
			if (field("B4")) {
				field("B4").disabled = true;
			}
			if (form("formname")) {
				submitting = true;
				form("formname").action = "BooksEditPopupUpdate.asp";
				form("formname").submit();
			}
			return true;
		};

		window.CheckSubmit = function () {
			var toSelect = field("selToAccHead");
			var selected;
			var response;
			if (valueOf("hToAccHead") !== "Y" || !toSelect) {
				return false;
			}
			selected = selectedValuesFrom(toSelect).join(",");
			response = postText("ContraEntryPopupUpdate.asp?OrgCode=" + encodeURIComponent(valueOf("hOrgCode")) + "&FromHead=" + encodeURIComponent(valueOf("hFromHead")) + "&ToHead=" + encodeURIComponent(selected));
			if (trim(response) !== "") {
				alert(response);
				return false;
			}
			setValue("hAction", "Done");
			returnAndClose("Done");
			return true;
		};

		window.DelMapBook = function () {
			var selected = checkedContraValues();
			var response;
			if (!selected.length) {
				alert("Select Mapped Book to Delete");
				return false;
			}
			response = postText("ContraEntryPopupDelete.asp?OrgCode=" + encodeURIComponent(valueOf("hOrgCode")) + "&FromHead=" + encodeURIComponent(valueOf("hFromHead")) + "&ToHead=" + encodeURIComponent(selected.join(",")));
			if (trim(response) !== "") {
				alert(response);
			} else {
				refreshContraBookRows();
			}
			return true;
		};

		window.ClearTable = function () {
			var table = byId("tblMap");
			if (!table) {
				return;
			}
			while (table.rows.length > 1) {
				table.deleteRow(1);
			}
		};

		window.CreateXML = function () {
			var xhr = new XMLHttpRequest();
			var text;
			xhr.open("GET", "../../Admin/Master/XMLGetNoSeriesPattern.asp", false);
			xhr.send(null);
			text = xhr.responseXML && xhr.responseXML.documentElement ? new XMLSerializer().serializeToString(xhr.responseXML) : xhr.responseText;
			if (trim(text) !== "") {
				loadXmlIntoIsland("SeriesNoData", text);
			} else if (trim(xhr.responseText)) {
				alert(xhr.responseText);
			}
			return true;
		};

		window.addEventListener("beforeunload", function () {
			if (!sent && !submitting) {
				returnValue(valueOf("hAction"));
			}
		});
	}

	function installBooksCreationEntry() {
		function selectedSeriesNode() {
			var root = xmlRoot("SeriesNoData");
			var seriesNo = valueOf("selNoSeries");
			return childElements(root, "Series").filter(function (node) {
				return attrAt(node, 0) === seriesNo;
			})[0] || null;
		}

		function periodText(seriesNode, entryNode) {
			var period = attrAt(entryNode, 1);
			var counterType = attrAt(seriesNode, 3);
			if (counterType === "M") {
				return "Month-" + period;
			}
			if (counterType === "Q") {
				return "Quater-" + period;
			}
			if (counterType === "Y") {
				return "Yearly";
			}
			return period;
		}

		function addTextCell(row, name, value, size, maxLength) {
			var input = document.createElement("input");
			input.type = "text";
			input.name = name || "";
			input.value = value == null ? "" : String(value);
			input.className = "Formelem";
			if (size) {
				input.size = size;
			}
			if (maxLength) {
				input.maxLength = maxLength;
			}
			addTableCell(row, "ExcelInputCell", "", input);
		}

		function setSeriesMetadata(seriesNode, entries) {
			if (!seriesNode) {
				setValue("hSeriesType", "");
				setValue("hSeriesLen", "");
				return;
			}
			if (entries.length === 12) {
				setValue("hSeriesType", "M");
			} else if (entries.length === 4) {
				setValue("hSeriesType", "Q");
			} else if (entries.length === 1) {
				setValue("hSeriesType", "Y");
			} else {
				setValue("hSeriesType", attrAt(seriesNode, 2));
			}
			setValue("hSeriesLen", attrAt(seriesNode, 4));
		}

		function clearBookTable(flag) {
			var table = byId("tblBook");
			var row;
			if (!table) {
				return;
			}
			while (table.rows.length) {
				table.deleteRow(0);
			}
			row = table.insertRow(0);
			addTableCell(row, "ExcelSerial", "Center", "S.No");
			addTableCell(row, "ExcelHeaderCell", "left", "Period");
			if (flag === "Y") {
				addTableCell(row, "ExcelHeaderCell", "left", "CR StartNo");
				addTableCell(row, "ExcelHeaderCell", "left", "CR Prefix");
				addTableCell(row, "ExcelHeaderCell", "left", "CR Suffix");
				addTableCell(row, "ExcelHeaderCell", "left", "DR StartNo");
				addTableCell(row, "ExcelHeaderCell", "left", "DR Prefix");
				addTableCell(row, "ExcelHeaderCell", "left", "DR Suffix");
			} else {
				addTableCell(row, "ExcelHeaderCell", "left", "StartNo");
				addTableCell(row, "ExcelHeaderCell", "left", "Prefix");
				addTableCell(row, "ExcelHeaderCell", "left", "Suffix");
			}
		}

		function xmlResponseText(xhr) {
			if (xhr.responseXML && xhr.responseXML.documentElement) {
				return new XMLSerializer().serializeToString(xhr.responseXML);
			}
			return xhr.responseText || "";
		}

		window.popUnitBooks = function () {
			var xhr = new XMLHttpRequest();
			var xmlText;
			var response;
			xhr.open("GET", "XMLGetDayBooks.asp", false);
			xhr.send(null);
			xmlText = xmlResponseText(xhr);
			if (trim(xmlText) === "" || trim(xmlText).charAt(0) !== "<") {
				if (trim(xhr.responseText)) {
					alert(xhr.responseText);
				}
				return false;
			}
			loadXmlIntoIsland("UnitBook", xmlText);
			response = postText("XMLSaveParty.asp?Name=Unit&Mod=Book", serializeXml("UnitBook"));
			if (trim(response) !== "") {
				alert(response);
				return false;
			}
			return true;
		};

		window.setPayRec = function () {
			var dayBook = valueOf("selDayBook");
			var select = field("selPayRecNo");
			if (!select) {
				return false;
			}
			select.options.length = 0;
			if (dayBook === "01" || dayBook === "02") {
				select.options[0] = new Option("Yes", "Y");
				select.options[1] = new Option("No", "N");
			} else {
				select.options[0] = new Option("No", "N");
			}
			return true;
		};

		window.DisplayBook = function () {
			var seriesNode = selectedSeriesNode();
			var entries;
			var index = 1;
			var separate = valueOf("selPayRecNo");
			clearBookTable(separate);
			if (!seriesNode || (field("selNoSeries") && field("selNoSeries").selectedIndex === 0)) {
				setSeriesMetadata(null, []);
				return false;
			}
			entries = childElements(seriesNode, "Entry");
			setSeriesMetadata(seriesNode, entries);
			entries.forEach(function (entry) {
				var row = byId("tblBook").insertRow(byId("tblBook").rows.length);
				var entryNo = attrAt(entry, 0);
				var startNo = attrAt(entry, 2);
				addTableCell(row, "ExcelSerial", "Center", String(index));
				addTableCell(row, "ExcelDisplayCell", "left", periodText(seriesNode, entry));
				if (separate === "Y") {
					addTextCell(row, "txtCrStartNo" + entryNo, startNo, 5, 4);
					addTextCell(row, "txtCrPrefix" + entryNo, attrAt(entry, 3), 12, 11);
					addTextCell(row, "txtCrSuffix" + entryNo, attrAt(entry, 4), 11, 10);
					addTextCell(row, "txtDrStartNo" + entryNo, startNo, 5, 4);
					addTextCell(row, "txtDrPrefix" + entryNo, attrAt(entry, 3), 12, 11);
					addTextCell(row, "txtDrSuffix" + entryNo, attrAt(entry, 4), 11, 10);
				} else {
					addTextCell(row, "txtStartNo" + entryNo, startNo, 5, 4);
					addTextCell(row, "txtPrefix" + entryNo, attrAt(entry, 3), 12, 11);
					addTextCell(row, "txtSuffix" + entryNo, attrAt(entry, 4), 11, 10);
				}
				index += 1;
			});
			return true;
		};

		window.popDayBookList = function () {
			var unitId = valueOf("hUnitId") || valueOf("hUnitID");
			openModernDialog(
				"PopBookNarrationList.asp?orgid=" + encodeURIComponent(unitId) + "&Mod=H",
				"",
				"dialogHeight:450px;dialogWidth:380px;center:Yes;help:No;resizable:No;status:No"
			);
			return true;
		};

		window.popSeriesNo = function () {
			var root = xmlRoot("SeriesNoData");
			var select = field("selNoSeries");
			if (!root || !select) {
				return false;
			}
			select.options.length = 0;
			select.options[0] = new Option("Select Number Series", "0");
			childElements(root).forEach(function (node) {
				if (attrAt(node, 3) === "M") {
					select.options[select.options.length] = new Option(attrAt(node, 1), attrAt(node, 0));
				}
			});
			return true;
		};

		window.ClearTable = clearBookTable;

		window.validateForm = function () {
			var seriesNode;
			var entries;
			if (field("selDayBook") && field("selDayBook").selectedIndex === 0) {
				alert("Select Day Book");
				field("selDayBook").focus();
				return false;
			}
			if (trim(valueOf("txtName")) === "") {
				alert("Enter Book Name");
				if (field("txtName")) {
					field("txtName").focus();
				}
				return false;
			}
			if (field("selNoSeries") && field("selNoSeries").selectedIndex === 0) {
				alert("Select No Series");
				field("selNoSeries").focus();
				return false;
			}
			if (trim(valueOf("hSeriesType")) === "") {
				seriesNode = selectedSeriesNode();
				entries = childElements(seriesNode, "Entry");
				setSeriesMetadata(seriesNode, entries);
			}
			if (field("B2")) {
				field("B2").disabled = true;
			}
			if (field("B4")) {
				field("B4").disabled = true;
			}
			if (form("formname")) {
				form("formname").action = "BooksCreationUpdate.asp";
				form("formname").submit();
			}
			return true;
		};

		window.CheckNumberSerious = function () {
			var xhr = new XMLHttpRequest();
			var xmlText;
			xhr.open("GET", "../../Admin/Master/XMLGetNoSeriesPattern.asp", false);
			xhr.send(null);
			xmlText = xmlResponseText(xhr);
			if (trim(xmlText) !== "" && trim(xmlText).charAt(0) === "<") {
				loadXmlIntoIsland("SeriesNoData", xmlText);
			} else if (trim(xhr.responseText)) {
				alert(xhr.responseText);
			}
			return true;
		};
	}

	function installGlAccountHeadDetails() {
		function selectedRadioValue(name, fallback) {
			var selected = fields(name).filter(function (radio) { return radio.checked; })[0];
			return selected ? selected.value : fallback || "";
		}

		function removeChildElements(root, name) {
			childElements(root, name).forEach(function (node) {
				root.removeChild(node);
			});
		}

		function childText(root, name, value, attrs) {
			var node = createXmlElement("GLHeadData", name);
			Object.keys(attrs || {}).forEach(function (key) {
				node.setAttribute(key, attrs[key]);
			});
			node.textContent = value == null ? "" : String(value);
			root.appendChild(node);
			return node;
		}

		function updateHeadNode(root, name, value, attrs) {
			removeChildElements(root, name);
			return childText(root, name, value, attrs);
		}

		function appendFlag(parent, name, flag) {
			var node = createXmlElement("GLHeadData", name);
			node.setAttribute("Flag", flag || "0");
			parent.appendChild(node);
			return node;
		}

		function checkedUnitIds() {
			return String(valueOf("hOrgid") || "").split(",").filter(function (unit) {
				return trim(unit) !== "";
			}).filter(function (unit) {
				var check = field("chkUnit" + unit);
				return check && check.checked;
			});
		}

		function unitName(unit) {
			var check = field("chkUnit" + unit);
			var parts = check ? String(check.value || "").split(":") : [];
			return valueOf("hUnitNameZ" + unit) || parts[1] || "";
		}

		function rebuildUnits(root) {
			var units = createXmlElement("GLHeadData", "Units");
			removeChildElements(root, "Units");
			root.appendChild(units);
			checkedUnitIds().forEach(function (unit) {
				var node = createXmlElement("GLHeadData", "UN");
				var summaryBook = createXmlElement("GLHeadData", "SummaryPostBook");
				node.setAttribute("Code", unit);
				node.setAttribute("Name", unitName(unit));
				node.setAttribute("OpBalance", valueOf("txtOpenBal" + unit));
				node.setAttribute("OpCRDR", selectedRadioValue("optOpenCD" + unit, "C"));
				units.appendChild(node);
				appendFlag(node, "IUT", selectedRadioValue("optIUT" + unit, "0"));
				appendFlag(node, "SummaryPosting", selectedRadioValue("optSummary" + unit, "0"));
				appendFlag(node, "SubLedger", selectedRadioValue("optSubledger" + unit, "0"));
				appendFlag(node, "CostCenter", selectedRadioValue("optCCZ" + unit, "0"));
				appendFlag(node, "Analytical", selectedRadioValue("optAnalZ" + unit, "0"));
				appendFlag(node, "Contra", selectedRadioValue("optContra" + unit, "0"));
				appendFlag(node, "TDS", selectedRadioValue("optTDS" + unit, "0"));
				appendFlag(node, "Memorandum", selectedRadioValue("optMemo" + unit, "0"));
				appendFlag(node, "CashTrans", selectedRadioValue("optTrans" + unit, "W"));
				summaryBook.setAttribute("BookCodes", valueOf("hSummAppSel" + unit));
				node.appendChild(summaryBook);
			});
		}

		function splitCodes(value) {
			return String(value || "").split(":").filter(function (item) {
				return trim(item) !== "";
			});
		}

		function splitNames(value) {
			return String(value || "").split(",").map(function (item) {
				return trim(item);
			}).filter(function (item) {
				return item !== "";
			});
		}

		function rebuildApplicationsAndBooks(root) {
			var apps = splitCodes(valueOf("hAppCode"));
			var appNames = splitNames(valueOf("txtAppUsed"));
			var books = splitCodes(valueOf("hBookCode"));
			var bookNames = splitNames(valueOf("txtBooks"));
			var node;
			removeChildElements(root, "Applications");
			removeChildElements(root, "Books");
			if (apps.length) {
				node = createXmlElement("GLHeadData", "Applications");
				root.appendChild(node);
				apps.forEach(function (code, index) {
					var app = createXmlElement("GLHeadData", "APP");
					app.setAttribute("Code", code);
					app.textContent = appNames[index] || "";
					node.appendChild(app);
				});
			}
			if (books.length) {
				node = createXmlElement("GLHeadData", "Books");
				node.setAttribute("Count", String(books.length));
				root.appendChild(node);
				books.forEach(function (code, index) {
					var book = createXmlElement("GLHeadData", "BK");
					book.setAttribute("Code", code);
					book.setAttribute("Name", bookNames[index] || "");
					node.appendChild(book);
				});
			}
		}

		function mergePartyData(root) {
			var partyRoot = xmlRoot("PartyData");
			var unitsNode = childElements(root, "Units")[0];
			if (!partyRoot || !unitsNode) {
				return;
			}
			childElements(unitsNode, "UN").forEach(function (unitNode) {
				var code = getAttribute(unitNode, "Code");
				childElements(partyRoot, "Unit").forEach(function (partyUnit) {
					if (getAttribute(partyUnit, "Code") === code) {
						childElements(partyUnit).forEach(function (parType) {
							unitNode.appendChild(parType.cloneNode(true));
						});
					}
				});
			});
		}

		function syncCurrentGlHeadXml(includePartyData) {
			var root = xmlRoot("GLHeadData") || loadXmlIntoIsland("GLHeadData", "<Root/>");
			updateHeadNode(root, "OpeningMonthYear", valueOf("hOpenYear"));
			updateHeadNode(root, "ClosingMonthYear", valueOf("hCloseYear"));
			updateHeadNode(root, "GroupCode", valueOf("hGCode"), { Name: valueOf("hGName") });
			updateHeadNode(root, "Description", valueOf("txtAccname"));
			updateHeadNode(root, "ShortName", valueOf("txtAccShortName"));
			updateHeadNode(root, "AccHeadNo", valueOf("hAccCode"));
			rebuildUnits(root);
			rebuildApplicationsAndBooks(root);
			if (includePartyData) {
				mergePartyData(root);
			}
			return root;
		}

		function saveCurrentGlHeadXml(includePartyData) {
			var root = syncCurrentGlHeadXml(includePartyData);
			var response = postText("XMLSaveParty.asp?Name=GLAccount&Mod=Head", new XMLSerializer().serializeToString(root.ownerDocument || root));
			if (trim(response) !== "") {
				alert(response);
				return false;
			}
			return true;
		}

		function replaceXmlIslandFromValue(name, value) {
			var text = "";
			if (!value) {
				return;
			}
			if (typeof value === "string") {
				text = value;
			} else if (value.nodeType === 9) {
				text = new XMLSerializer().serializeToString(value);
			} else if (value.ownerDocument) {
				text = new XMLSerializer().serializeToString(value);
			} else if (value.documentElement) {
				text = new XMLSerializer().serializeToString(value.documentElement);
			}
			if (trim(text)) {
				loadXmlIntoIsland(name, text);
			}
		}

		function selectedGlRoot() {
			return xmlObject("GLHeadData");
		}

		function removeUnitScopedChildren(containerName, unit) {
			var root = xmlRoot("GLHeadData");
			if (!root) {
				return;
			}
			childElements(root, containerName).forEach(function (container) {
				childElements(container).forEach(function (item) {
					if (getAttribute(item, "UNCode") === unit) {
						container.removeChild(item);
					}
				});
			});
		}

		function updateEntryIcon(id, enabled) {
			var img = byId(id);
			if (img) {
				img.disabled = !enabled;
			}
		}

		window.CheckDetails = function (unit, checkbox, accountHead) {
			var response;
			if (!checkbox || checkbox.checked || !confirm("Do you want to delete this unit Information?")) {
				return true;
			}
			response = postText("GLHeadUnitDelete.asp?AccHead=" + encodeURIComponent(accountHead) + "&UnitCode=" + encodeURIComponent(unit));
			if (trim(response) !== "") {
				alert(response);
			} else if (form("formname")) {
				form("formname").submit();
			}
			return true;
		};

		window.ValidatePopAnalHead = function (radio) {
			var unit = String(radio && radio.name || "").split("Z")[1] || "";
			updateEntryIcon("imgAnalyticalEntryZ" + unit, radio && radio.value === "1");
			if (radio && radio.value !== "1") {
				removeUnitScopedChildren("Analytical", unit);
			}
			return true;
		};

		window.ValidateCostCenterHead = function (radio) {
			var unit = String(radio && radio.name || "").split("Z")[1] || "";
			updateEntryIcon("imgCostCenterZ" + unit, radio && radio.value === "1");
			if (radio && radio.value !== "1") {
				removeUnitScopedChildren("CostCenter", unit);
			}
			return true;
		};

		window.SelectAccHead = function () {
			openModernDialog("comAccountGroupTreePopup.asp", "", "dialogHeight:510px;dialogWidth:350px;Status:No;Help:No", function (returnValue) {
				var parts;
				if (trim(returnValue) === "") {
					return;
				}
				parts = String(returnValue).split(":");
				setValue("hGCode", String(parts[0] || "").substring(2));
				setValue("hGName", parts.slice(1).join(":"));
				if (byId("AccGroupName")) {
					byId("AccGroupName").textContent = valueOf("hGName");
				}
			});
			return true;
		};

		window.ControlAccount = function (unit) {
			var check = field("chkUnit" + unit);
			var subledger = selectedRadioValue("optSubledger" + unit, "0");
			var unitNameValue = unitName(unit);
			if (!check || !check.checked) {
				alert("Select the Applicable Units");
				return false;
			}
			if (subledger === "1") {
				if (!saveCurrentGlHeadXml(false)) {
					return false;
				}
				openModernDialog("GLHeadParSubTypePopup.asp?UnitName=" + encodeURIComponent(unitNameValue), "", "dialogWidth:600px;dialogHeight:485px;Status:No;Help:No;", function (value) {
					replaceXmlIslandFromValue("PartyData", value);
				});
			}
			return true;
		};

		window.PopCostCenter = function (unit) {
			var headName = valueOf("hAccName") || valueOf("txtAccname");
			var query = "AccHead=" + encodeURIComponent(valueOf("hAccCode")) +
				"&HeadName=" + encodeURIComponent(headName) +
				"&GroupName=" + encodeURIComponent(valueOf("hGName")) +
				"&Units=" + encodeURIComponent(unit) +
				"&hSelCostCode=" + encodeURIComponent(valueOf("hSelCostCodeZ" + unit));
			openModernDialog("GLHeadCostCenterPopup.asp?" + query, selectedGlRoot(), "dialogHeight:350px;dialogWidth:400px;center:Yes;help:No;resizable:No;status:No", function (value) {
				replaceXmlIslandFromValue("GLHeadData", value);
			});
			return true;
		};

		window.PopAnalyticalHead = function (unit) {
			var headName = valueOf("hAccName") || valueOf("txtAccname");
			var query = "AccHead=" + encodeURIComponent(valueOf("hAccCode")) +
				"&HeadName=" + encodeURIComponent(headName) +
				"&GroupName=" + encodeURIComponent(valueOf("hGName")) +
				"&Units=" + encodeURIComponent(unit) +
				"&hSelAnayCode=" + encodeURIComponent(valueOf("hSelAnayCodeZ" + unit));
			openModernDialog("GLHeadAnalyticalPopup.asp?" + query, selectedGlRoot(), "dialogHeight:350px;dialogWidth:400px;center:Yes;help:No;resizable:No;status:No", function (value) {
				replaceXmlIslandFromValue("GLHeadData", value);
			});
			return true;
		};

		window.PopUsed = function (kind) {
			var selected = kind === "A" ? valueOf("txtAppUsed") : valueOf("txtBooks");
			openModernDialog("glAppandBooksUsed.asp?sTempValues=" + encodeURIComponent(kind + "?" + selected), "", "dialogHeight:350px;dialogWidth:350px;center:Yes;help:No;resizable:No;status:No", function (result) {
				var parts;
				if (String(result || "0/0") === "0/0") {
					return;
				}
				parts = String(result || "").split("/");
				if (kind === "A") {
					setValue("hAppCode", parts[0] || "");
					setValue("txtAppUsed", String(parts[1] || "").replace(/^,/, ""));
				} else {
					setValue("hBookCode", parts[0] || "");
					setValue("txtBooks", String(parts[1] || "").replace(/^,/, ""));
				}
			});
			return true;
		};

		window.SelVouType = function (radios, unit, accountHead) {
			var summary = fields("optSummary" + unit);
			var app = field("hSummAppSel" + unit);
			var selected = "A?" + (app ? app.value : "") + "?" + trim(unit) + "?" + trim(accountHead);
			if (summary[0] && summary[0].checked) {
				openModernDialog("GlSummVouTy.asp?sTempValues=" + encodeURIComponent(selected), "", "dialogHeight:290px;dialogWidth:270px;center:Yes;help:No;resizable:No;status:No", function (result) {
					if (String(result || "0/0") !== "0/0" && app) {
						app.value = result || "";
					}
				});
			}
			return true;
		};

		window.Finaldone = function () {
			if (trim(valueOf("hGCode")) === "") {
				alert("Select Account Group Name");
				if (field("txtAccname")) {
					field("txtAccname").focus();
				}
				return false;
			}
			if (trim(valueOf("txtAccname")) === "") {
				alert("Enter Account Description");
				if (field("txtAccname")) {
					field("txtAccname").focus();
				}
				return false;
			}
			if (trim(valueOf("txtAccShortName")) === "") {
				alert("Enter Account Short Description");
				if (field("txtAccShortName")) {
					field("txtAccShortName").focus();
				}
				return false;
			}
			if (!checkedUnitIds().length) {
				alert("Select any one Applicable Units");
				return false;
			}
			if (!saveCurrentGlHeadXml(true)) {
				return false;
			}
			if (form("formname")) {
				form("formname").action = "GLCreate_Edit_Update.asp?Acion=" + encodeURIComponent(valueOf("hAction"));
				form("formname").submit();
			}
			return true;
		};

		window.init = function () {
			var root = xmlRoot("GLHeadData");
			if (!root) {
				return;
			}
			childElements(childElements(root, "Units")[0], "UN").forEach(function (unit) {
				var code = getAttribute(unit, "Code");
				childElements(unit).forEach(function (flag) {
					if (flag.nodeName === "CostCenter") {
						updateEntryIcon("imgCostCenterZ" + code, getAttribute(flag, "Flag") === "1");
					}
					if (flag.nodeName === "Analytical") {
						updateEntryIcon("imgAnalyticalEntryZ" + code, getAttribute(flag, "Flag") === "1");
					}
				});
			});
		};
	}

	function installGlHeadScopedSelectionPopup(kind) {
		var containerName = kind === "cost" ? "CostCenter" : "Analytical";
		var itemName = kind === "cost" ? "CC" : "AN";
		var fieldPrefix = kind === "cost" ? "chkCostCenter" : "chkAnalyticalZ";
		var totalField = kind === "cost" ? "hTotalCostCenter" : "hTotalAnalCode";

		function sourceRoot() {
			return xmlRoot(ensureDialogArgs());
		}

		function itemValue(node) {
			if (kind === "cost") {
				return getAttribute(node, "UNCode") + ":" + getAttribute(node, "GRCode") + ":" + getAttribute(node, "CCCode");
			}
			return getAttribute(node, "UNCode") + ":" + getAttribute(node, "Code") + ":" + getAttribute(node, "GRCode");
		}

		window.init = function () {
			var root = sourceRoot();
			if (!root) {
				return;
			}
			childElements(root, containerName).forEach(function (container) {
				childElements(container).forEach(function (item) {
					var value = itemValue(item);
					for (var index = 1; index <= toNumber(valueOf(totalField)); index += 1) {
						if (field(fieldPrefix + index) && field(fieldPrefix + index).value === value) {
							field(fieldPrefix + index).checked = true;
							break;
						}
					}
				});
			});
		};

		window.CheckSubmit = function () {
			var root = sourceRoot();
			var doc = root && root.ownerDocument;
			var container;
			if (!root || !doc) {
				window.close();
				return false;
			}
			container = childElements(root, containerName)[0];
			if (!container) {
				container = doc.createElement(containerName);
				root.appendChild(container);
			}
			childElements(container).forEach(function (item) {
				if (getAttribute(item, "UNCode") === valueOf("hUnitCode")) {
					container.removeChild(item);
				}
			});
			for (var index = 1; index <= toNumber(valueOf(totalField)); index += 1) {
				var checkbox = field(fieldPrefix + index);
				var parts;
				var node;
				if (checkbox && checkbox.checked) {
					parts = String(checkbox.value || "").split(":");
					node = doc.createElement(itemName);
					node.setAttribute("UNCode", parts[0] || "");
					if (kind === "cost") {
						node.setAttribute("GRCode", parts[1] || "");
						node.setAttribute("CCCode", parts[2] || "");
					} else {
						node.setAttribute("Code", parts[1] || "");
						node.setAttribute("GRCode", parts[2] || "");
					}
					container.appendChild(node);
				}
			}
			returnAndClose(root);
			return true;
		};

		window.Popup_Close = function () {
			returnAndClose(sourceRoot());
			return true;
		};
	}

	function installGlHeadPartySubTypePopup() {
		function createSubTypeRoot() {
			return xmlRoot("SubTypeData") || loadXmlIntoIsland("SubTypeData", '<Root action="Cancel"/>');
		}

		function sourceRoot() {
			return xmlRoot("GLHeadData") || xmlRoot(ensureDialogArgs());
		}

		window.NewSubType = function () {
			openModernDialog("ParTypeCreatEntryPopup.asp", "", "", function (value) {
				if (value === "Done" && form("formname")) {
					form("formname").submit();
				}
			});
			return true;
		};

		window.PageSubmit = function () {
			var glRoot = sourceRoot();
			var root = createSubTypeRoot();
			var doc = xmlDocument("SubTypeData") || root.ownerDocument;
			clearChildren(root);
			childElements(childElements(glRoot, "Units")[0], "UN").forEach(function (unit) {
				var unitCode = getAttribute(unit, "Code");
				var unitNode = doc.createElement("Unit");
				unitNode.setAttribute("Code", unitCode);
				root.appendChild(unitNode);
				for (var index = 1; index <= toNumber(valueOf("hSubTypeCount")); index += 1) {
					var checkbox = field("selPartyTypeZ" + index);
					var parts;
					var parType;
					if (checkbox && checkbox.checked) {
						parts = String(checkbox.value || "").split("?");
						if (trim(parts[0]) === trim(unitCode)) {
							parType = doc.createElement("ParType");
							parType.setAttribute("UnitId", parts[0] || "");
							parType.setAttribute("PartyType", parts[1] || "");
							parType.setAttribute("PartySubType", parts[2] || "");
							unitNode.appendChild(parType);
						}
					}
				}
			});
			root.setAttribute("Action", "Done");
			returnAndClose(root);
			return true;
		};
	}

	function installGlAccountHeadGrid() {
		function setCheckboxFromHidden(checkName, hiddenName) {
			var check = field(checkName);
			if (check) {
				check.checked = valueOf(hiddenName) === "1";
			}
		}

		function checkedValue(checkName) {
			var check = field(checkName);
			return check && check.checked ? "1" : "0";
		}

		function submitTo(url) {
			if (form("formname")) {
				form("formname").action = url;
				form("formname").submit();
			}
		}

		window.init = function () {
			setCheckboxFromHidden("chkContra", "hContra");
			setCheckboxFromHidden("chkParty", "hSubLed");
			setCheckboxFromHidden("chkTDS", "hTDS");
			setCheckboxFromHidden("chkSummary", "hSumPosting");
		};

		window.CheckSubmit = function () {
			setValue("hContra", checkedValue("chkContra"));
			setValue("hSubLed", checkedValue("chkParty"));
			setValue("hTDS", checkedValue("chkTDS"));
			setValue("hSumPosting", checkedValue("chkSummary"));
			if (form("formname")) {
				form("formname").submit();
			}
			return true;
		};

		window.GetAccHead = function (select) {
			var xhr = new XMLHttpRequest();
			var target = field("selGLGroup");
			var root;
			xhr.open("GET", "GetAccHeadForGLGroup.asp?GCode=" + encodeURIComponent(select && select.value || ""), false);
			xhr.send(null);
			if (xhr.responseXML && xhr.responseXML.documentElement) {
				root = xhr.responseXML.documentElement;
				loadXmlIntoIsland("AccHeadData", new XMLSerializer().serializeToString(xhr.responseXML));
			} else if (trim(xhr.responseText)) {
				root = loadXmlIntoIsland("AccHeadData", xhr.responseText);
			}
			if (!root) {
				if (trim(xhr.responseText)) {
					alert(xhr.responseText);
				}
				return false;
			}
			if (target) {
				target.options.length = 1;
				childElements(root).forEach(function (node) {
					target.options[target.options.length] = new Option(getAttribute(node, "Name"), getAttribute(node, "No"));
				});
			}
			return true;
		};

		window.CreateNewHead = function () {
			submitTo("GLCreate_Edit_AccHeadDet.asp");
			return true;
		};

		window.AccountsCategory = function (index) {
			var radios = fields("radAccHead");
			var radio = radios[toNumber(index)] || radios[0];
			if (radio) {
				submitTo("GLACCHEADGRID.ASP?Category=" + encodeURIComponent(radio.value));
			}
			return true;
		};

		window.EditAccHead = function (unitId, accountNo, accountName, groupCode, groupName) {
			setValue("selUnitId", unitId);
			setValue("hHeadValue", accountNo);
			setValue("hHeadName", accountName);
			setValue("GCode", groupCode);
			setValue("GName", groupName);
			submitTo("GLCreate_Edit_AccHeadDet.asp");
			return true;
		};
	}

	function selectedValuesFrom(select) {
		return Array.prototype.slice.call(select && select.options || []).filter(function (option) {
			return option.selected;
		}).map(function (option) {
			return option.value;
		});
	}

	function postText(url, body) {
		var xhr = new XMLHttpRequest();
		xhr.open("POST", url, false);
		xhr.send(body == null ? null : body);
		return xhr.responseText || "";
	}

	function checkedContraValues() {
		return Array.prototype.slice.call(document.querySelectorAll('input[type="checkbox"]')).filter(function (checkbox) {
			return /^chkbox/i.test(checkbox.name || "") && checkbox.checked && !checkbox.disabled;
		}).map(function (checkbox) {
			return checkbox.value;
		});
	}

	function renderContraRows(root, includeFromColumn, fromName) {
		var tbl = byId("tblMap");
		var headerRows = includeFromColumn ? 2 : 1;
		var index = 1;
		if (!tbl) {
			return;
		}
		while (tbl.rows.length > headerRows) {
			tbl.deleteRow(headerRows);
		}
		childElements(root).forEach(function (node) {
			var row = tbl.insertRow(tbl.rows.length);
			var checkbox = document.createElement("input");
			checkbox.type = "checkbox";
			checkbox.name = "chkBox" + index;
			checkbox.className = "FormElem";
			checkbox.value = getAttribute(node, "No");
			checkbox.disabled = getAttribute(node, "Records") === "Y";
			addTableCell(row, "ExcelHeaderCell", "left", String(index));
			addTableCell(row, "ExcelDisplayCell", "center", checkbox);
			if (includeFromColumn) {
				addTableCell(row, "ExcelDisplayCell", "left", fromName || "");
			}
			addTableCell(row, "ExcelDisplayCell", "left", getAttribute(node, "Name"));
			index += 1;
		});
		setValue("hRowCnt", index - 1);
	}

	function refreshContraMappings(fromHead, includeFromColumn) {
		var xhr = new XMLHttpRequest();
		var root;
		var fromSelect = field("selFormHead");
		var fromName = fromSelect ? selectedText(fromSelect) : "";
		xhr.open("GET", "GetMappedContraDetails.asp?OrgCode=" + encodeURIComponent(valueOf("hOrgCode")) + "&FromHead=" + encodeURIComponent(fromHead), false);
		xhr.send(null);
		root = xhr.responseXML && xhr.responseXML.documentElement;
		if (!root && trim(xhr.responseText)) {
			root = new DOMParser().parseFromString(xhr.responseText, "text/xml").documentElement;
		}
		if (!root) {
			if (trim(xhr.responseText)) {
				alert(xhr.responseText);
			}
			return;
		}
		loadXmlIntoIsland("AccData", new XMLSerializer().serializeToString(root.ownerDocument || root));
		renderContraRows(root, includeFromColumn, fromName);
	}

	function installContraList() {
		window.CheckSubmit = function () {
			var toSelect = field("selToAccHead");
			var selected;
			var response;
			if (valueOf("hToAccHead") !== "Y" || !toSelect) {
				return false;
			}
			selected = selectedValuesFrom(toSelect).join(",");
			response = postText("ContraEntryPopupUpdate.asp?OrgCode=" + encodeURIComponent(valueOf("hOrgCode")) + "&FromHead=" + encodeURIComponent(valueOf("hFromHead")) + "&ToHead=" + encodeURIComponent(selected));
			if (trim(response) !== "") {
				alert(response);
			} else {
				returnAndClose("Done");
			}
			return true;
		};
		window.DelMapBook = function () {
			var selected = checkedContraValues();
			var response;
			if (!selected.length) {
				alert("Select Mapped Book to Delete");
				return false;
			}
			response = postText("ContraEntryPopupDelete.asp?OrgCode=" + encodeURIComponent(valueOf("hOrgCode")) + "&FromHead=" + encodeURIComponent(valueOf("hFromHead")) + "&ToHead=" + encodeURIComponent(selected.join(",")));
			if (trim(response) !== "") {
				alert(response);
			} else {
				refreshContraMappings(valueOf("hFromHead"), false);
			}
			return true;
		};
	}

	function installContraConfiguration() {
		window.DisplayBook = function () {
			var xhr = new XMLHttpRequest();
			var root;
			xhr.open("GET", "XMLOrgGLHead.asp?orgID=" + encodeURIComponent(valueOf("hOrgCode")), false);
			xhr.send(null);
			root = xhr.responseXML && xhr.responseXML.documentElement;
			if (!root && trim(xhr.responseText)) {
				root = new DOMParser().parseFromString(xhr.responseText, "text/xml").documentElement;
			}
			if (root) {
				loadXmlIntoIsland("BookData", new XMLSerializer().serializeToString(root.ownerDocument || root));
				window.popFromHead();
			}
		};
		window.popFromHead = function () {
			var root = xmlRoot("BookData");
			var fromSelect = field("selFormHead");
			var toSelect = field("selToHead");
			if (!root || !fromSelect || !toSelect) {
				return;
			}
			fromSelect.options.length = 0;
			toSelect.options.length = 0;
			childElements(root, "Head").forEach(function (node) {
				fromSelect.options[fromSelect.options.length] = new Option(getAttribute(node, "AccountDesc"), getAttribute(node, "AccountNo"));
			});
			window.popToHead();
		};
		window.popToHead = function () {
			var root = xmlRoot("BookData");
			var fromSelect = field("selFormHead");
			var toSelect = field("selToHead");
			if (!root || !fromSelect || !toSelect) {
				return;
			}
			toSelect.options.length = 0;
			childElements(root, "Head").forEach(function (node) {
				if (getAttribute(node, "AccountNo") === selectedValue(fromSelect)) {
					childElements(node, "ToHead").forEach(function (child) {
						toSelect.options[toSelect.options.length] = new Option(getAttribute(child, "AccountDesc"), getAttribute(child, "AccountNo"));
					});
				}
			});
		};
		window.CheckSubmit = function () {
			var fromSelect = field("selFormHead");
			var toSelect = field("selToHead");
			var fromHead = selectedValue(fromSelect);
			var selected;
			var response;
			if (!fromHead) {
				alert("select From Book");
				if (fromSelect) {
					fromSelect.focus();
				}
				return false;
			}
			selected = selectedValuesFrom(toSelect);
			if (!selected.length) {
				alert("Select To Book");
				if (toSelect) {
					toSelect.focus();
				}
				return false;
			}
			response = postText("ContraEntryPopupUpdate.asp?OrgCode=" + encodeURIComponent(valueOf("hOrgCode")) + "&FromHead=" + encodeURIComponent(fromHead) + "&ToHead=" + encodeURIComponent(selected.join(",")));
			if (trim(response) !== "") {
				alert(response);
			} else if (form("formname")) {
				form("formname").submit();
			}
			return true;
		};
		window.DelMapBook = function () {
			var fromSelect = field("selFormHead");
			var fromHead = selectedValue(fromSelect) || valueOf("hFromHead");
			var selected = checkedContraValues();
			var response;
			if (!selected.length) {
				alert("Select Mapped Book to Delete");
				return false;
			}
			response = postText("ContraEntryPopupDelete.asp?OrgCode=" + encodeURIComponent(valueOf("hOrgCode")) + "&FromHead=" + encodeURIComponent(fromHead) + "&ToHead=" + encodeURIComponent(selected.join(",")));
			if (trim(response) !== "") {
				alert(response);
			} else {
				refreshContraMappings(fromHead, true);
			}
			return true;
		};
	}

	function setButtonDisabled(name, disabled) {
		var item = field(name);
		if (item) {
			item.disabled = !!disabled;
		}
	}

	function saveXmlIsland(url, xmlName) {
		var xhr = new XMLHttpRequest();
		xhr.open("POST", url, false);
		try {
			xhr.setRequestHeader("Content-Type", "text/xml");
		} catch (ignore) {}
		xhr.send(serializeXml(xmlName));
		return xhr.responseText || "";
	}

	function nextXmlNumber(root, nodeName) {
		var max = 0;
		childElements(root, nodeName).forEach(function (node) {
			max = Math.max(max, toNumber(getAttribute(node, "No")));
		});
		return max + 1;
	}

	function installPartyContactPopup() {
		function root() {
			return xmlRoot("OutData");
		}

		function resetForm() {
			["txtContact", "txtDesignation", "txtMailId", "txtName"].forEach(function (name) {
				setValue(name, "");
			});
		}

		function buttonsEditing(editing) {
			setButtonDisabled("btnNext", editing);
			setButtonDisabled("btnAdd", editing);
			setButtonDisabled("btnUpdate", !editing);
			setButtonDisabled("btnDel", !editing);
		}

		function addOrUpdateNode() {
			var docRoot = root();
			var no = valueOf("hEntNo");
			var node;
			if (!docRoot) {
				return null;
			}
			if (!no) {
				no = String(nextXmlNumber(docRoot, "Entry"));
				setValue("hEntNo", no);
			}
			node = childElements(docRoot, "Entry").filter(function (item) {
				return getAttribute(item, "No") === no;
			})[0];
			if (!node) {
				node = createXmlElement("OutData", "Entry");
				docRoot.appendChild(node);
			}
			node.setAttribute("No", no);
			node.setAttribute("Name", valueOf("txtName"));
			node.setAttribute("Desig", valueOf("txtDesignation"));
			node.setAttribute("PersonFor", valueOf("txtContact"));
			node.setAttribute("Maillid", valueOf("txtMailId"));
			setValue("hRecCount", Math.max(toNumber(valueOf("hRecCount")), toNumber(no) + 1));
			return node;
		}

		window.ClearTable = function () {
			var tbl = byId("tblBin");
			var row;
			if (!tbl) {
				return;
			}
			while (tbl.rows.length) {
				tbl.deleteRow(0);
			}
			row = tbl.insertRow(0);
			addTableCell(row, "ExcelHeaderCell", "center", "S.No.");
			addTableCell(row, "ExcelHeaderCell", "center", " ");
			addTableCell(row, "ExcelHeaderCell", "center", "Name");
			addTableCell(row, "ExcelHeaderCell", "center", "Designation");
			addTableCell(row, "ExcelHeaderCell", "center", "Contact For");
			addTableCell(row, "ExcelHeaderCell", "center", "Email ID");
		};
		window.popDisplayTable = function () {
			var tbl = byId("tblBin");
			var rowNo = 1;
			window.ClearTable();
			if (!tbl || !root()) {
				return;
			}
			childElements(root(), "Entry").forEach(function (node) {
				var row = tbl.insertRow(tbl.rows.length);
				addTableCell(row, "ExcelDisplayCell", "center", String(rowNo));
				addTableCell(row, "ExcelDisplayCell", "center", "<a href=\"javascript:EditEntry('" + getAttribute(node, "No") + "')\" class=\"ExcelDisplayLink\">Edit</a>");
				addTableCell(row, "ExcelDisplayCell", "left", getAttribute(node, "Name"));
				addTableCell(row, "ExcelDisplayCell", "left", getAttribute(node, "Desig"));
				addTableCell(row, "ExcelDisplayCell", "left", getAttribute(node, "PersonFor"));
				addTableCell(row, "ExcelDisplayCell", "left", getAttribute(node, "Maillid"));
				rowNo += 1;
			});
		};
		window.PopulateContact = function () {
			window.popDisplayTable();
		};
		window.Form_Reset = resetForm;
		window.addDataNode = addOrUpdateNode;
		window.addEntry = function (flag) {
			if (flag === "A") {
				if (window.validate && !window.validate()) {
					return false;
				}
				setValue("hEntNo", "");
				addOrUpdateNode();
				window.popDisplayTable();
				resetForm();
				return true;
			}
			if (trim(valueOf("txtName")) !== "") {
				if (window.validate && !window.validate()) {
					return false;
				}
				setValue("hEntNo", "");
				addOrUpdateNode();
			}
			return window.SaveXML();
		};
		window.EditEntry = function (no) {
			var node = childElements(root(), "Entry").filter(function (item) {
				return getAttribute(item, "No") === String(no);
			})[0];
			if (!node) {
				return;
			}
			setValue("hEntNo", no);
			setValue("txtName", getAttribute(node, "Name"));
			setValue("txtDesignation", getAttribute(node, "Desig"));
			setValue("txtContact", getAttribute(node, "PersonFor"));
			setValue("txtMailId", getAttribute(node, "Maillid"));
			buttonsEditing(true);
		};
		window.updateEntry = function () {
			if (window.validate && !window.validate()) {
				return false;
			}
			addOrUpdateNode();
			window.popDisplayTable();
			resetForm();
			buttonsEditing(false);
			return true;
		};
		window.DelEntry = function () {
			var docRoot = root();
			var no = valueOf("hEntNo");
			var node = childElements(docRoot, "Entry").filter(function (item) {
				return getAttribute(item, "No") === no;
			})[0];
			if (node) {
				docRoot.removeChild(node);
			}
			window.popDisplayTable();
			resetForm();
			buttonsEditing(false);
		};
		window.clearXML = function () {
			clearChildren(root());
			window.popDisplayTable();
			resetForm();
			setValue("hEntNo", "");
		};
		window.SaveXML = function () {
			var response = saveXmlIsland("ParContactPopupUpdate.asp?PartyCode=" + encodeURIComponent(valueOf("hPartyCode")), "OutData");
			if (trim(response) !== "") {
				alert(response);
			} else {
				returnAndClose("Done");
			}
			return true;
		};
	}

	function installPartyLocationPopup() {
		function root() {
			return xmlRoot("OutData");
		}

		function resetForm() {
			[
				"txtAddress1",
				"txtAddress2",
				"txtCity",
				"txtCountry",
				"txtEccNo",
				"txtLocationName",
				"txtPanNo",
				"txtSalesCenteral",
				"txtSalesLocal",
				"txtState"
			].forEach(function (name) {
				setValue(name, "");
			});
		}

		function buttonsEditing(editing) {
			setButtonDisabled("btnAdd", editing);
			setButtonDisabled("btnNext", editing);
			setButtonDisabled("btnUpdate", !editing);
			setButtonDisabled("btnupdate", !editing);
			setButtonDisabled("btnDel", !editing);
		}

		function activeLocationNodes() {
			return childElements(root(), "Loc").filter(function (node) {
				return getAttribute(node, "Status") !== "2";
			});
		}

		function addOrUpdateNode() {
			var docRoot = root();
			var no = valueOf("hEntNo");
			var node;
			if (!docRoot) {
				return null;
			}
			if (!no) {
				no = String(nextXmlNumber(docRoot, "Loc"));
				setValue("hEntNo", no);
			}
			node = childElements(docRoot, "Loc").filter(function (item) {
				return getAttribute(item, "No") === no;
			})[0];
			if (!node) {
				node = createXmlElement("OutData", "Loc");
				docRoot.appendChild(node);
			}
			node.setAttribute("No", no);
			node.setAttribute("Name", valueOf("txtLocationName"));
			node.setAttribute("Address1", valueOf("txtAddress1"));
			node.setAttribute("Address2", valueOf("txtAddress2"));
			node.setAttribute("City", valueOf("txtCity"));
			node.setAttribute("State", valueOf("txtState"));
			node.setAttribute("Country", valueOf("txtCountry"));
			node.setAttribute("ECCNo", valueOf("txtEccNo"));
			node.setAttribute("SalesLocal", valueOf("txtSalesLocal"));
			node.setAttribute("SalesCentral", valueOf("txtSalesCenteral"));
			node.setAttribute("PANNo", valueOf("txtPanNo"));
			node.setAttribute("Status", getAttribute(node, "Status") || "1");
			setValue("hRecCount", Math.max(toNumber(valueOf("hRecCount")), toNumber(no) + 1));
			return node;
		}

		window.ClearTable = function () {
			var tbl = byId("tblBin");
			var row;
			if (!tbl) {
				return;
			}
			while (tbl.rows.length) {
				tbl.deleteRow(0);
			}
			row = tbl.insertRow(0);
			addTableCell(row, "ExcelHeaderCell", "center", "S.No.");
			addTableCell(row, "ExcelHeaderCell", "center", " ");
			addTableCell(row, "ExcelHeaderCell", "center", "Location Name");
			addTableCell(row, "ExcelHeaderCell", "center", "Address");
			addTableCell(row, "ExcelHeaderCell", "center", "");
		};
		window.popDisplayTable = function () {
			var tbl = byId("tblBin");
			var rowNo = 1;
			window.ClearTable();
			if (!tbl || !root()) {
				return;
			}
			activeLocationNodes().forEach(function (node) {
				var row = tbl.insertRow(tbl.rows.length);
				var address = [getAttribute(node, "Address1"), getAttribute(node, "Address2"), getAttribute(node, "City"), getAttribute(node, "State"), getAttribute(node, "Country")].filter(Boolean).join("<br>");
				var tax = "ECCNo: " + getAttribute(node, "ECCNo") + "<br>" +
					"Local Sale Tax: " + getAttribute(node, "SalesLocal") + "<br>" +
					"Central Sales Tax:" + getAttribute(node, "SalesCentral") + "<br>" +
					"IT PanNo :" + getAttribute(node, "PANNo") + "<br>";
				addTableCell(row, "ExcelDisplayCell", "center", String(rowNo));
				addTableCell(row, "ExcelDisplayCell", "center", "<a href=\"javascript:EditEntry('" + getAttribute(node, "No") + "')\" class=\"ExcelDisplayCell\"><b>Edit</b></a>");
				addTableCell(row, "ExcelDisplayCell", "left", getAttribute(node, "Name"));
				addTableCell(row, "ExcelDisplayCell", "left", address);
				addTableCell(row, "ExcelDisplayCell", "left", tax);
				rowNo += 1;
			});
		};
		window.PopulateLoc = function () {
			window.popDisplayTable();
		};
		window.Form_Reset = resetForm;
		window.addDataNode = addOrUpdateNode;
		window.addEntry = function (flag) {
			if (flag === "A") {
				if (window.validate && !window.validate()) {
					return false;
				}
				setValue("hEntNo", "");
				addOrUpdateNode();
				window.popDisplayTable();
				resetForm();
				return true;
			}
			if (trim(valueOf("txtLocationName")) !== "") {
				if (window.validate && !window.validate()) {
					return false;
				}
				setValue("hEntNo", "");
				addOrUpdateNode();
			}
			return window.SaveXML();
		};
		window.EditEntry = function (no) {
			var node = childElements(root(), "Loc").filter(function (item) {
				return getAttribute(item, "No") === String(no);
			})[0];
			if (!node) {
				return;
			}
			setValue("hEntNo", no);
			setValue("txtLocationName", getAttribute(node, "Name"));
			setValue("txtAddress1", getAttribute(node, "Address1"));
			setValue("txtAddress2", getAttribute(node, "Address2"));
			setValue("txtCity", getAttribute(node, "City"));
			setValue("txtState", getAttribute(node, "State"));
			setValue("txtCountry", getAttribute(node, "Country"));
			setValue("txtEccNo", getAttribute(node, "ECCNo"));
			setValue("txtSalesLocal", getAttribute(node, "SalesLocal"));
			setValue("txtSalesCenteral", getAttribute(node, "SalesCentral"));
			setValue("txtPanNo", getAttribute(node, "PANNo"));
			buttonsEditing(true);
		};
		window.updateEntry = function () {
			if (window.validate && !window.validate()) {
				return false;
			}
			addOrUpdateNode();
			window.popDisplayTable();
			resetForm();
			buttonsEditing(false);
			return true;
		};
		window.DispLocDet = function () {
			window.EditEntry(valueOf("selLocName"));
		};
		window.DelEntry = function () {
			var no = valueOf("hEntNo");
			var node = childElements(root(), "Loc").filter(function (item) {
				return getAttribute(item, "No") === no;
			})[0];
			if (node) {
				if (getAttribute(node, "Status") === "0" || getAttribute(node, "Status") === "E") {
					alert("Party Location is Present in Sale Transaction Not able to Delete ");
				} else {
					root().removeChild(node);
				}
			}
			window.popDisplayTable();
			resetForm();
			buttonsEditing(false);
		};
		window.clearXML = function () {
			clearChildren(root());
			window.popDisplayTable();
			resetForm();
			setValue("hEntNo", "");
		};
		window.SaveXML = function () {
			var response = saveXmlIsland("ParLocationPopupUpdate.asp?PartyCode=" + encodeURIComponent(valueOf("hPartyCode")), "OutData");
			if (trim(response) !== "") {
				alert(response);
			} else {
				returnAndClose("Done");
			}
			return true;
		};
	}

	function installPartyCreateEditModals() {
		var detailDialogFeatures = "dialogHeight:440px;dialogWidth:600px;center:Yes;help:No;resizable:No;status:No";

		function submitForm() {
			var frm = form("formname");
			if (frm) {
				frm.submit();
			}
		}

		function currentPartyCode(partyCode) {
			return trim(partyCode) || trim(valueOf("hPartyCode"));
		}

		function requireSaved(message) {
			if (trim(valueOf("hPartyCode")) === "") {
				alert(message);
				return false;
			}
			return true;
		}

		function openPartyDialog(page, partyCode, features, message) {
			if (!requireSaved(message)) {
				return false;
			}
			openModernDialog(page + "?PartyCode=" + encodeURIComponent(currentPartyCode(partyCode)), "", features || detailDialogFeatures, function (value) {
				if (trim(value) === "Done") {
					submitForm();
				}
			});
			return true;
		}

		function focusField(name) {
			var item = field(name);
			if (item && item.focus) {
				item.focus();
			}
			return item;
		}

		function setReadOnly(name, readOnly) {
			var item = field(name);
			if (item) {
				item.readOnly = !!readOnly;
			}
		}

		function checkedValue(name, fallback) {
			var checked = fields(name).filter(function (item) { return item.checked; })[0];
			return checked ? checked.value : fallback || "";
		}

		function ensureRoot(xmlName, rootName) {
			var root = xmlRoot(xmlName);
			if (!root) {
				loadXmlIntoIsland(xmlName, "<" + rootName + "/>");
				root = xmlRoot(xmlName);
			}
			return root;
		}

		function ensureChild(root, name) {
			var node = firstElement(root, name);
			if (!node) {
				node = createXmlElement("PartyData", name);
				root.appendChild(node);
			}
			return node;
		}

		function setTextNode(root, name, value) {
			ensureChild(root, name).textContent = value == null ? "" : String(value);
		}

		function selectedUnitPairs() {
			var codes = trim(valueOf("hUnitCode")) ? valueOf("hUnitCode").split(":") : [];
			var names = trim(valueOf("hUnitName")) ? valueOf("hUnitName").split(":") : [];
			var ownUnit = trim(valueOf("hOwnUnit"));
			return codes.map(function (code, index) {
				return {
					code: trim(code),
					name: names[index] || ""
				};
			}).filter(function (unit) {
				return unit.code && unit.code !== ownUnit;
			});
		}

		function updateUnits(root, rebuild) {
			var units = ensureChild(root, "Units");
			if (!rebuild && childElements(units, "UN").length) {
				return units;
			}
			clearChildren(units);
			selectedUnitPairs().forEach(function (unit) {
				var node = createXmlElement("PartyData", "UN");
				node.setAttribute("Code", unit.code);
				node.setAttribute("Name", unit.name);
				units.appendChild(node);
			});
			return units;
		}

		function collectPartyXml(rebuildUnits) {
			var root = ensureRoot("PartyData", "Root");
			var groupFlag = field("chkGroupCompany") && field("chkGroupCompany").checked ? valueOf("chkGroupCompany") : "0";
			var groupType = groupFlag === "0" ? "N" : checkedValue("radGroupType", "N");
			var units = firstElement(root, "Units");
			var shouldRebuildUnits = rebuildUnits || !units || !childElements(units, "UN").length;
			setTextNode(root, "ParCode", valueOf("hPartyCode"));
			setTextNode(root, "ParName", valueOf("txtPartyName"));
			setTextNode(root, "ShortName", valueOf("txtShortName"));
			setTextNode(root, "Address1", valueOf("txtAddress1"));
			setTextNode(root, "Address2", valueOf("txtAddress2"));
			setTextNode(root, "PinCode", valueOf("txtPinCode"));
			setTextNode(root, "City", valueOf("txtCity"));
			setTextNode(root, "State", valueOf("txtState"));
			setTextNode(root, "Country", valueOf("txtCountry"));
			setTextNode(root, "Phone", valueOf("txtPhone"));
			setTextNode(root, "Mobile", valueOf("txtMobileNo"));
			setTextNode(root, "Fax", valueOf("txtFax"));
			setTextNode(root, "Email", valueOf("txtEmail"));
			setTextNode(root, "Website", valueOf("txtWebsite"));
			setTextNode(root, "ECCNo", valueOf("txtECCNo"));
			setTextNode(root, "PANNo", valueOf("txtPanNo"));
			setTextNode(root, "CreatedBy", valueOf("hCreatedBy"));
			setTextNode(root, "OwnUnit", valueOf("hOwnUnit"));
			setTextNode(root, "TINNumber", valueOf("txtTinNo"));
			setAttribute(ensureChild(root, "Sales"), "Local", valueOf("txtSalesLocal"));
			setAttribute(ensureChild(root, "Sales"), "Central", valueOf("txtSalesCentral"));
			setAttribute(ensureChild(root, "Group"), "Flag", groupFlag);
			setAttribute(ensureChild(root, "Group"), "Type", groupType);
			setAttribute(ensureChild(root, "Group"), "ParentCompany", valueOf("hParentPartyCode"));
			setAttribute(ensureChild(root, "Active"), "Flag", valueOf("hInActive"));
			updateUnits(root, shouldRebuildUnits);
			return root;
		}

		function postPartyXml() {
			var response = postText("XMLSaveParty.asp?Name=Party&Mod=Master", serializeXml("PartyData"));
			if (trim(response) !== "") {
				alert(response);
			}
		}

		function appendImported(parent, node) {
			var doc = xmlDocument("PartyData");
			parent.appendChild(doc && doc.importNode ? doc.importNode(node, true) : node.cloneNode(true));
		}

		function mergeExistingPartyDetails() {
			var xhr;
			var tempRoot;
			var root;
			var units;
			if (trim(valueOf("hAction")).toUpperCase() !== "EDIT" || trim(valueOf("hParUnit")).toUpperCase() !== "N") {
				return;
			}
			xhr = new XMLHttpRequest();
			xhr.open("GET", "ParGetUnitDetails.asp", false);
			xhr.send(null);
			if (xhr.responseXML && xhr.responseXML.documentElement) {
				loadXmlIntoIsland("TempData", new XMLSerializer().serializeToString(xhr.responseXML));
			} else if (trim(xhr.responseText) !== "") {
				loadXmlIntoIsland("TempData", xhr.responseText);
			}
			tempRoot = xmlRoot("TempData");
			root = ensureRoot("PartyData", "Root");
			units = ensureChild(root, "Units");
			if (!tempRoot) {
				if (trim(xhr.responseText) !== "") {
					alert(xhr.responseText);
				}
				return;
			}
			childElements(tempRoot).forEach(function (node) {
				var unit;
				if (String(node.nodeName).toLowerCase() === "partytype") {
					unit = childElements(units, "UN").filter(function (unitNode) {
						return getAttribute(unitNode, "Code") === getAttribute(node, "Unit");
					})[0];
					if (unit) {
						appendImported(unit, node);
					}
				} else if (String(node.nodeName).toLowerCase() === "agent") {
					appendImported(root, node);
				}
			});
		}

		function partyDataHasTypes() {
			var units = firstElement(ensureRoot("PartyData", "Root"), "Units");
			var unitNodes = childElements(units, "UN");
			return unitNodes.length > 0 && unitNodes.every(function (unit) {
				return childElements(unit).length > 0;
			});
		}

		function clearPartyText() {
			[
				"txtPartyName",
				"txtAddress1",
				"txtAddress2",
				"txtCity",
				"txtPinCode",
				"txtState",
				"txtCountry",
				"txtPhone",
				"txtFax",
				"txtEmail",
				"txtWebsite"
			].forEach(function (name) {
				setValue(name, "");
			});
			setValue("hOwnUnit", "");
		}

		function applyOwnUnit(unitId, unitName) {
			var node = childElements(xmlRoot("UNITDET"), "UNIT").filter(function (item) {
				return getAttribute(item, "UnitID") === unitId;
			})[0];
			setValue("hOwnUnit", unitId);
			setValue("txtPartyName", (valueOf("hOrgName") || valueOf("horgName")) + " - " + unitName);
			setReadOnly("txtPartyName", true);
			if (node) {
				setValue("txtAddress1", getAttribute(node, "Add1"));
				setValue("txtAddress2", getAttribute(node, "Add2"));
				setValue("txtCity", getAttribute(node, "City"));
				setValue("txtPinCode", getAttribute(node, "PostCode"));
				setValue("txtState", getAttribute(node, "State"));
				setValue("txtCountry", getAttribute(node, "Country"));
				setValue("txtPhone", getAttribute(node, "Phone"));
				setValue("txtFax", getAttribute(node, "Fax"));
				setValue("txtEmail", getAttribute(node, "EmailID"));
				setValue("txtWebsite", getAttribute(node, "Web"));
			}
		}

		function unitCheckbox(index) {
			return field("chkUnitZ" + index);
		}

		function unitCodeFromCheckbox(checkbox) {
			return trim(String(checkbox && checkbox.value || "").split(":")[0]);
		}

		function checkPartyUnits(unitText) {
			var units = trim(unitText) ? String(unitText).split(":").filter(function (item) { return trim(item) !== ""; }) : [];
			var rowCount = toNumber(valueOf("hUnitrow"));
			var allUnits = field("chkUnitZ0");
			var index;
			if (allUnits) {
				allUnits.checked = units.length > 0 && units.length === rowCount;
			}
			for (index = 1; index <= rowCount; index += 1) {
				if (unitCheckbox(index)) {
					unitCheckbox(index).checked = units.indexOf(unitCodeFromCheckbox(unitCheckbox(index))) !== -1;
				}
			}
		}

		function disableUsedUnits(typeText) {
			var usedUnits = trim(typeText) ? String(typeText).split(":").filter(function (item) { return trim(item) !== "" && trim(item) !== "0"; }) : [];
			var rowCount = toNumber(valueOf("hUnitrow"));
			var index;
			if (!usedUnits.length) {
				return;
			}
			for (index = 1; index <= rowCount; index += 1) {
				if (unitCheckbox(index) && unitCheckbox(index).checked && usedUnits.indexOf(unitCodeFromCheckbox(unitCheckbox(index))) !== -1) {
					unitCheckbox(index).disabled = true;
				}
			}
			if (rowCount === 0 && field("chkUnitZ0") && field("chkUnitZ0").checked) {
				field("chkUnitZ0").disabled = true;
			}
		}

		function applyGroupType(groupType) {
			var groupCheck = field("chkGroupCompany");
			var radios = fields("radGroupType");
			var index;
			if (["P", "C", "B"].indexOf(String(groupType)) === -1) {
				return;
			}
			if (groupCheck) {
				groupCheck.checked = true;
			}
			for (index = 0; index < radios.length; index += 1) {
				radios[index].disabled = false;
				radios[index].checked = radios[index].value === groupType;
			}
		}

		function populatePartyNode(node) {
			var units = getAttribute(node, "Units");
			var partyGroupType = getAttribute(node, "PartyGroupCoyType");
			var transactionUnits = getAttribute(node, "InTrans");
			setValue("txtShortName", getAttribute(node, "OrgnPartyCode"));
			setValue("txtAddress1", getAttribute(node, "AddressLine1"));
			setValue("txtAddress2", getAttribute(node, "AddressLine2"));
			setValue("txtCity", getAttribute(node, "City"));
			setValue("txtState", getAttribute(node, "State"));
			setValue("txtCountry", getAttribute(node, "Country"));
			setValue("txtPhone", getAttribute(node, "PhoneNos"));
			setValue("txtMobileNo", getAttribute(node, "MobileNos"));
			setValue("txtFax", getAttribute(node, "FaxNos"));
			setValue("txtEmail", getAttribute(node, "Email"));
			setValue("txtWebsite", getAttribute(node, "WebsiteURL"));
			setValue("txtPinCode", getAttribute(node, "Pincode").replace(/\s+/g, ""));
			setValue("txtECCNo", getAttribute(node, "ExciseControlCode"));
			setValue("txtSalesLocal", getAttribute(node, "LocalSTNoandDT"));
			setValue("txtSalesCentral", getAttribute(node, "CentralSTNoandDT"));
			setValue("txtPanNo", getAttribute(node, "IncomeTaxPANNo"));
			setValue("txtPartyName", getAttribute(node, "PartyName"));
			setValue("txtTinNo", getAttribute(node, "TINNumber"));
			if (getAttribute(node, "Useable") === "1" && field("chkActive")) {
				field("chkActive").checked = true;
			}
			checkPartyUnits(units);
			disableUsedUnits(transactionUnits);
			applyGroupType(partyGroupType);
			window.CheckUnit();
		}

		window.ViewData = function () {
			var frm = form("formname");
			if (!requireSaved("Party Details Cannot view because Party is not available")) {
				return false;
			}
			frm.action = "ParDetailsView.asp?PartyCode=" + encodeURIComponent(valueOf("hPartyCode"));
			frm.submit();
			return true;
		};
		window.ControlData = function () {
			var frm = form("formname");
			if (!requireSaved("Party Controls Cannot View because Party is not available")) {
				return false;
			}
			frm.action = "PartyControlData.asp?PartyCode=" + encodeURIComponent(valueOf("hPartyCode"));
			frm.submit();
			return true;
		};
		window.GoToMain = function () {
			var frm = form("formname");
			if (frm) {
				frm.action = "ParDisplayGrid.asp";
				frm.submit();
			}
			return true;
		};
		window.GetGroup = function () {
			var selected = fields("radGroupType").filter(function (radio) { return radio.checked; })[0];
			if (!selected || selected.value === "P") {
				setValue("hParentPartyCode", "0");
				setValue("ParentPartyName", "");
			}
			return true;
		};
		window.popPartyDet = function (partyCode) {
			var xhr;
			var root;
			if (trim(partyCode) === "") {
				return true;
			}
			xhr = new XMLHttpRequest();
			xhr.open("GET", "XMLGetPartyDet.asp?PartyCode=" + encodeURIComponent(partyCode), false);
			xhr.send(null);
			if (xhr.responseXML && xhr.responseXML.documentElement) {
				loadXmlIntoIsland("OutData", new XMLSerializer().serializeToString(xhr.responseXML));
			} else if (trim(xhr.responseText) !== "") {
				loadXmlIntoIsland("OutData", xhr.responseText);
			}
			root = xmlRoot("OutData");
			childElements(root).forEach(populatePartyNode);
			return true;
		};
		window.Fun_Contact = function (partyCode) {
			return openPartyDialog("ParContactPopup.asp", partyCode, detailDialogFeatures, "Save the Basic Party Details before entering Contact Details");
		};
		window.Fun_Location = function (partyCode) {
			return openPartyDialog("ParLocationPopup.asp", partyCode, detailDialogFeatures, "Save the Basic Party Details before entering Location Details");
		};
		window.Fun_Preference = function (partyCode) {
			return openPartyDialog("ParPerferencePopup.asp", partyCode, detailDialogFeatures, "Save the Basic Party Details before entering Preference Details");
		};
		window.Fun_Agent = function (partyCode) {
			return openPartyDialog("ParAgentSelectPopup.asp", partyCode, detailDialogFeatures, "Save the Basic Party Details before entering Agent Details");
		};
		window.Fun_Rep = function (partyCode) {
			return openPartyDialog("RepSelectionEntry.asp", partyCode, "dialogHeight:250px;dialogWidth:400px;center:Yes;help:No;resizable:No;status:No", "Save the Basic Party Details before entering Rep. Details");
		};
		window.Cleartxt = clearPartyText;
		window.GetUnit = function () {
			var checkbox = field("ChkOwnUnit");
			if (checkbox && checkbox.checked) {
				openModernDialog("ParCreationUnitSelPopup.asp", "", "dialogHeight:240px;dialogWidth:400px;center:Yes;help:No;resizable:No;status:No", function (value) {
					var parts;
					if (trim(value) !== "") {
						parts = String(value).split(":");
						applyOwnUnit(parts[0], parts.slice(1).join(":"));
					} else {
						setReadOnly("txtPartyName", false);
						clearPartyText();
					}
				});
			} else {
				setReadOnly("txtPartyName", false);
				clearPartyText();
			}
			return true;
		};
		window.CheckUnit = function () {
			var all = field("chkUnitZ0");
			var rowCount = toNumber(valueOf("hUnitrow"));
			var unitIds = [];
			var unitNames = [];
			var index;
			function addUnit(checkbox) {
				var parts;
				if (!checkbox) {
					return;
				}
				parts = String(checkbox.value || "").split(":");
				if (trim(parts[0]) && trim(parts[0]) !== "0") {
					unitIds.push(trim(parts[0]));
					unitNames.push(parts.slice(1).join(":"));
				}
			}
			if (all && all.checked) {
				for (index = 1; index <= rowCount; index += 1) {
					addUnit(field("chkUnitZ" + index));
				}
			} else {
				for (index = 1; index <= rowCount; index += 1) {
					if (field("chkUnitZ" + index) && field("chkUnitZ" + index).checked) {
						addUnit(field("chkUnitZ" + index));
					}
				}
			}
			setValue("hUnitCode", unitIds.join(":"));
			setValue("hUnitName", unitNames.join(":"));
			setValue("hInActive", field("chkActive") && field("chkActive").checked ? "1" : "0");
			return true;
		};
		window.CheckForm = function () {
			if (trim(valueOf("txtShortName")) === "") {
				alert("Enter Party Code ");
				focusField("txtShortName");
				return false;
			}
			if (trim(valueOf("txtCity")) === "") {
				alert("Enter City ");
				focusField("txtCity");
				return false;
			}
			if (trim(valueOf("hUnitCode")) === "") {
				alert("Selec the Units");
				return false;
			}
			return true;
		};
		window.SaveXML = function () {
			collectPartyXml(true);
			postPartyXml();
			mergeExistingPartyDetails();
			postPartyXml();
			return true;
		};
		window.SaveXMLFinal = function () {
			collectPartyXml(trim(valueOf("hParUnit")).toUpperCase() !== "Y");
			postPartyXml();
			mergeExistingPartyDetails();
			postPartyXml();
			return true;
		};
		window.PopulatePartyTypes = function (partyCode, action) {
			window.CheckUnit();
			if (!window.CheckForm()) {
				return false;
			}
			window.SaveXML();
			openModernDialog("ParUnitPopup.asp?PartyCode=" + encodeURIComponent(currentPartyCode(partyCode)) + "&Action=" + encodeURIComponent(action || valueOf("hAction")), xmlObject("PartyData"), "dialogHeight:500px;dialogWidth:650px;center:Yes;help:No;resizable:No;status:No", function (value) {
				var xmlText;
				if (!value) {
					return;
				}
				xmlText = value.nodeType ? new XMLSerializer().serializeToString(value) : value.xml || value.XML || "";
				if (trim(xmlText) !== "") {
					setValue("hParUnit", "Y");
					loadXmlIntoIsland("PartyData", xmlText);
				}
			});
			return true;
		};
		window.PageSubmit = function () {
			var frm = form("formname");
			window.CheckUnit();
			if (!window.CheckForm()) {
				return false;
			}
			window.SaveXMLFinal();
			if (!partyDataHasTypes()) {
				alert("Enter Party Types for the Units");
				return false;
			}
			if (trim(valueOf("hUnitCode")) === "0") {
				alert("Select any Unit ");
				return false;
			}
			frm.action = trim(valueOf("hPartyCode")) !== ""
				? "ParCreate_Edit_EntryInsert.asp?Action=Edit&PartyCode=" + encodeURIComponent(valueOf("hPartyCode"))
				: "ParCreate_Edit_EntryInsert.asp?Action=Insert";
			setButtonDisabled("B2", true);
			setButtonDisabled("B3", true);
			frm.submit();
			return true;
		};
	}

	function installPartyUnitPopup() {
		function root() {
			return xmlRoot("PartyData");
		}

		function selectedRadioValue(name) {
			var checked = fields(name).filter(function (radio) { return radio.checked; })[0];
			return checked ? checked.value : "";
		}

		function removeExistingAgentNodes(docRoot) {
			childElements(docRoot, "Agent").forEach(function (node) {
				docRoot.removeChild(node);
			});
		}

		window.AddNew = function () {
			openModernDialog("ParTypeCreatEntryPopup.asp", "", "Status:No;Help:No;", function (value) {
				var frm;
				if (trim(value) === "Done") {
					frm = form("formname");
					if (frm) {
						frm.submit();
					}
				}
			});
			return true;
		};
		window.CheckPartyChk = function (checkValue, checkbox) {
			if (String(checkValue) === "A") {
				alert("Party Type has Transaction Entry Not able to Remove the Party Type Reference!! ");
				if (checkbox) {
					checkbox.checked = true;
				}
			}
		};
		window.PageSubmit = function () {
			var docRoot = root();
			var agentFlag = false;
			if (!docRoot) {
				return false;
			}
			removeExistingAgentNodes(docRoot);
			childElements(firstElement(docRoot, "Units"), "UN").forEach(function (unitNode) {
				var unitId = getAttribute(unitNode, "Code") || getAttribute(unitNode, "Unit");
				var rowCount = toNumber(valueOf("hRowCntUnitZ" + unitId));
				var index;
				clearChildren(unitNode);
				for (index = 1; index <= rowCount; index += 1) {
					var checkbox = field("chkParType" + unitId + "Z" + index);
					var parts;
					var node;
					if (!checkbox || !checkbox.checked) {
						continue;
					}
					parts = String(checkbox.value || "").split("?");
					node = createXmlElement("PartyData", "Partytype");
					node.setAttribute("Type", parts[0] || "");
					node.setAttribute("SubType", parts[1] || "");
					node.setAttribute("OpBalance", valueOf("txtBalance" + unitId + "Z" + (parts[0] || "") + "Z" + (parts[1] || "") + "Z" + (parts[2] || "")));
					node.setAttribute("OpCRDR", selectedRadioValue("radCRDR" + unitId + "Z" + (parts[0] || "") + "Z" + (parts[1] || "") + "Z" + (parts[2] || "")));
					node.setAttribute("OpeningMonthYear", valueOf("hFinFromYear"));
					node.setAttribute("ClosingMonthYear", valueOf("hFinToYear"));
					unitNode.appendChild(node);
					if (trim(parts[0]) === "CR" && toNumber(parts[1]) < 3) {
						agentFlag = true;
					}
				}
			});
			var agent = createXmlElement("PartyData", "Agent");
			agent.setAttribute("Flag", agentFlag ? "1" : "0");
			docRoot.appendChild(agent);
			returnAndClose(docRoot);
			return true;
		};
		window.window_onunload = function () {
			if (!sent) {
				returnValue(root());
			}
		};
		window.addEventListener("beforeunload", window.window_onunload);
	}

	function installAppBookUsed() {
		defaultSelectionReturn = "0/0";
		window.finaldone = function () {
			var values = [];
			var names = [];
			selectedChecklistItems().forEach(function (checkbox) {
				var parts = String(checkbox.value || "").split(":");
				values.push(parts[0] || "");
				names.push(parts[1] || "");
			});
			defaultSelectionReturn = values.length ? ":" + values.join(":") + "/" + "," + names.join(",") : "/";
			returnAndClose(defaultSelectionReturn);
		};
		window.finalcancel = function () {
			returnAndClose(defaultSelectionReturn);
		};
		window.window_onunload = function () {
			if (!sent) {
				returnValue(defaultSelectionReturn);
			}
		};
		window.CheckVal = function (selected) {
			var values = String(selected || "").split(",");
			fields("chkSelVal").forEach(function (checkbox) {
				var parts = String(checkbox.value || "").split(":");
				checkbox.checked = values.some(function (value) {
					return value.indexOf(parts[1] || "") >= 0;
				});
			});
		};
		installEscape(window.finalcancel);
	}

	function installVoucherTypeSelection() {
		defaultSelectionReturn = "0/0";
		window.finaldone = function () {
			var values = selectedChecklistItems().map(function (checkbox) {
				return checkbox.value;
			});
			defaultSelectionReturn = values.join(":");
			returnAndClose(defaultSelectionReturn);
		};
		window.finalcancel = function () {
			returnAndClose(defaultSelectionReturn);
		};
		window.window_onunload = function () {
			if (!sent) {
				returnValue(defaultSelectionReturn);
			}
		};
		window.CheckTrans = function () {
			var checks = fields("chkSelVal");
			[
				["hCaAnt", 0],
				["hBaAnt", 1],
				["hPurAnt", 2],
				["hSalAnt", 3],
				["hCreAnt", 4],
				["hDebAnt", 5],
				["hGJAnt", 6]
			].forEach(function (entry) {
				var check = checks[entry[1]];
				if (check && String(valueOf(entry[0])) !== "0" && check.checked) {
					check.disabled = true;
				}
			});
		};
		window.CheckVal = function (selected) {
			var values = String(selected || "").split(",");
			fields("chkSelVal").forEach(function (checkbox) {
				checkbox.checked = values.some(function (value) {
					return value.indexOf(checkbox.value) >= 0;
				});
			});
			window.CheckTrans();
		};
		installEscape(window.finalcancel);
	}

	function installPartySubtypeEntry() {
		installSimpleUnload("Cancel");
		window.CheckSubmit = function (action, parType, parSubType) {
			var frm = form("formname");
			var radios = fields("radParType");
			var selectedType = radios.filter(function (radio) { return radio.checked; })[0];
			var checked;
			var parts;
			if (!frm) {
				return false;
			}
			if (trim(action) === "C") {
				if (!selectedType) {
					alert("Select Party Type");
					if (radios[0] && radios[0].focus) {
						radios[0].focus();
					}
					return false;
				}
				if (trim(valueOf("txtSubTypeName")) === "") {
					alert("Enter Sub-Type Name");
					field("txtSubTypeName").select();
					return false;
				}
				if (trim(valueOf("txtSubTypeShortName")) === "") {
					alert("Enter Sub-Type Short Name");
					field("txtSubTypeShortName").select();
					return false;
				}
				frm.action = "ParTypeCreatUpdatePopup.asp?Action=" + action;
			} else if (trim(action) === "E") {
				frm.action = "ParTypeCreatEntryPopup.asp?Action=" + action + "&ParType=" + parType + "&ParSubType=" + parSubType;
			} else if (trim(action) === "U") {
				frm.action = "ParTypeCreatUpdatePopup.asp?Action=" + action + "&ParType=" + parType + "&ParSubType=" + parSubType;
			} else if (trim(action) === "D") {
				checked = fields("chkParType").concat(Array.prototype.slice.call(document.querySelectorAll('input[type="checkbox"][name^="chkParType"]:checked')));
				checked = checked.filter(function (checkbox, index, list) {
					return checkbox.checked && list.indexOf(checkbox) === index;
				});
				if (checked.length !== 1) {
					alert("Select any one Sub Type to Delete");
					return false;
				}
				parts = String(checked[0].value || "").split(":");
				frm.action = "ParTypeCreatUpdatePopup.asp?Action=" + action + "&ParType=" + (parts[0] || "") + "&ParSubType=" + (parts[1] || "");
			}
			if (field("B2")) {
				field("B2").disabled = true;
			}
			frm.submit();
			return true;
		};
	}

	function installAutoClose() {
		window.init = window.Init = function () {
			if (config.message) {
				alert(config.message);
			}
			returnAndClose(config.returnValue || "Done");
		};
		installSimpleUnload(config.returnValue || "Done");
	}

	function installMiscPartyCreate() {
		window.AddDetails = function () {
			var root = xmlRoot("Party");
			var xhr;
			if (!root) {
				return false;
			}
			[
				["PartyName", "txtName"],
				["PartyShortName", "txtShortName"],
				["Add1", "txtAddress1"],
				["Add2", "txtAddress2"],
				["City", "txtCity"],
				["Pin", "txtPinCode"],
				["State", "txtState"],
				["Country", "txtCountry"],
				["EMail", "txtEmail"],
				["ITPan", "txtPanNo"],
				["Phone", "txtPhone"],
				["Fax", "txtFax"],
				["Mobile", "txtMobileNo"],
				["Url", "txtWebsite"]
			].forEach(function (entry) {
				setAttribute(root, entry[0], valueOf(entry[1]));
			});
			xhr = new XMLHttpRequest();
			xhr.open("POST", "MsiParUpdate.asp?", false);
			xhr.send(serializeXml("Party"));
			if (trim(xhr.responseText) === "") {
				alert("Party Created ");
				returnAndClose(valueOf("txtName"));
			} else {
				alert(xhr.responseText);
			}
			return true;
		};
	}

	function populateRepAreas(selectedArea) {
		var select = field("SelRepArea");
		var xhr = new XMLHttpRequest();
		var root;
		if (!select) {
			return;
		}
		xhr.open("GET", "/Common/XMLGetRepresentingArea.asp", false);
		xhr.send(null);
		root = xhr.responseXML && xhr.responseXML.documentElement;
		if (!root && trim(xhr.responseText)) {
			root = new DOMParser().parseFromString(xhr.responseText, "text/xml").documentElement;
		}
		if (!root) {
			if (trim(xhr.responseText)) {
				alert(xhr.responseText);
			}
			return;
		}
		loadXmlIntoIsland("RepAreaData", new XMLSerializer().serializeToString(root.ownerDocument || root));
		select.options.length = 1;
		childElements(root).forEach(function (node) {
			var value = getAttribute(node, "ACode");
			var option = new Option(getAttribute(node, "AName"), value);
			select.options[select.options.length] = option;
			if (trim(selectedArea) === trim(value)) {
				option.selected = true;
			}
		});
	}

	function installRepresentativeSelection() {
		window.PopulateContPerson = function (obj) {
			var areaValue = obj && obj.value || "0";
			var select = field("SelContPerson");
			var xhr;
			var root;
			if (!select) {
				return;
			}
			select.options.length = 0;
			if (areaValue === "0") {
				select.options[0] = new Option("Name", "0");
				return;
			}
			xhr = new XMLHttpRequest();
			xhr.open("GET", "/Common/XMLGetConPersonForRepArea.asp?AreaID=" + encodeURIComponent(areaValue), false);
			xhr.send(null);
			root = xhr.responseXML && xhr.responseXML.documentElement;
			if (!root && trim(xhr.responseText)) {
				root = new DOMParser().parseFromString(xhr.responseText, "text/xml").documentElement;
			}
			if (!root) {
				if (trim(xhr.responseText)) {
					alert(xhr.responseText);
				}
				return;
			}
			loadXmlIntoIsland("ConPerForArea", new XMLSerializer().serializeToString(root.ownerDocument || root));
			if (childElements(root).length) {
				childElements(root).forEach(function (node) {
					select.options[select.options.length] = new Option(getAttribute(node, "CPName"), getAttribute(node, "ID") + "|" + getAttribute(node, "LCode"));
				});
			} else {
				select.options[0] = new Option("Not Available", "0|0");
			}
		};
		window.PopulateArea = function () {
			populateRepAreas(valueOf("hAreaCode"));
		};
		window.Init = function () {
			var area = field("SelRepArea");
			var contact = field("SelContPerson");
			var selectedAgent = trim(valueOf("hAgentEntryID"));
			window.PopulateArea();
			window.PopulateContPerson(area);
			if (contact) {
				Array.prototype.forEach.call(contact.options, function (option) {
					if (trim(String(option.value).split("|")[0]) === selectedAgent) {
						option.selected = true;
					}
				});
			}
		};
		window.CreateRep = function () {
			openModernDialog("RepCreationEntry.asp", "", "dialogHeight:440px;dialogWidth:600px;center:Yes;help:No;resizable:No;status:No", function () {
				window.PopulateArea();
			});
		};
		window.checkSubmit = function () {
			var area = field("SelRepArea");
			var contact = field("SelContPerson");
			var root = xmlRoot("AgentData");
			var node;
			var xhr;
			if (!area || selectedValue(area) === "0") {
				alert("Select Area Name");
				return false;
			}
			if (!contact || selectedValue(contact) === "0") {
				alert("Select Representative Name");
				return false;
			}
			clearChildren(root);
			node = createXmlElement("AgentData", "AGENT");
			node.setAttribute("AgentEntryID", trim(selectedValue(contact)).split("|")[0]);
			node.setAttribute("AreaCode", selectedValue(area));
			root.appendChild(node);
			xhr = new XMLHttpRequest();
			xhr.open("POST", "XMLSaveParty.asp?Name=RepAllocation&Mod=Party", false);
			xhr.send(serializeXml("AgentData"));
			xhr = new XMLHttpRequest();
			xhr.open("POST", "RepSelectionInsert.asp?PartyCode=" + encodeURIComponent(valueOf("hPartyCode")), false);
			xhr.send(null);
			if (trim(xhr.responseText) !== "") {
				alert(xhr.responseText);
			} else {
				returnAndClose("Done");
			}
			return true;
		};
	}

	function installRepresentativeCreation() {
		window.PopulateArea = function () {
			populateRepAreas(valueOf("hAreaCode"));
		};
		window.Init = function () {
			window.PopulateArea();
		};
		window.ChangeType = function () {
			var radios = fields("radIntExt");
			var combo = field("cmbAgentType");
			if (combo && radios[0] && radios[1]) {
				combo.disabled = radios[0].checked;
			}
		};
		window.AddArea = function () {
			openModernDialog("../../Sales/Master/SalRepAreaEntry.asp", "", "dialogWidth:400px;dialogHeight:300px;Status:No", function () {
				window.PopulateArea();
			});
		};
		window.checkSubmit = function () {
			var frm = form("formname");
			var root = xmlRoot("AgentData");
			var header;
			var contactDetails;
			var contact;
			var xhr;
			if (trim(valueOf("txtAgentName")) === "") {
				alert("Enter Rep. Name");
				field("txtAgentName").focus();
				return false;
			}
			if (trim(valueOf("txtSAgentName")) === "") {
				alert("Enter Rep. Short Name");
				field("txtSAgentName").focus();
				return false;
			}
			if (trim(valueOf("txtEmail")) !== "" && window.checkmailid && !window.checkmailid(trim(valueOf("txtEmail")))) {
				field("txtEmail").focus();
				return false;
			}
			clearChildren(root);
			header = createXmlElement("AgentData", "AGENTHEADER");
			[
				["AgentName", "txtAgentName"],
				["AgentSName", "txtSAgentName"],
				["Add1", "txtAddress1"],
				["Add2", "txtAddress2"],
				["City", "txtCity"],
				["PinCode", "txtPin"],
				["Phone", "txtPhone"],
				["Fax", "txtFax"],
				["Mobile", "txtMobile"],
				["Email", "txtEmail"]
			].forEach(function (entry) {
				header.setAttribute(entry[0], valueOf(entry[1]));
			});
			header.setAttribute("State", "");
			header.setAttribute("Country", "");
			header.setAttribute("URL", "");
			header.setAttribute("AgentType", "");
			header.setAttribute("IntOrExt", "I");
			root.appendChild(header);
			contactDetails = createXmlElement("AgentData", "CONTACTDETAIL");
			root.appendChild(contactDetails);
			contact = createXmlElement("AgentData", "Contact");
			contact.setAttribute("Name", valueOf("txtAgentName"));
			contact.setAttribute("Desig", "");
			contact.setAttribute("ContactFor", "");
			contact.setAttribute("Email", valueOf("txtEmail"));
			contact.setAttribute("Address1", valueOf("txtAddress1"));
			contact.setAttribute("Address2", valueOf("txtAddress2"));
			contact.setAttribute("City", valueOf("txtCity"));
			contact.setAttribute("State", "");
			contact.setAttribute("Country", "");
			contact.setAttribute("RepArea", selectedValue(field("SelRepArea")));
			contactDetails.appendChild(contact);
			xhr = new XMLHttpRequest();
			xhr.open("POST", "XMLSaveParty.asp?Name=Rep&Mod=Party", false);
			xhr.send(serializeXml("AgentData"));
			frm.action = "RepCreationInsert.asp?PartyCode=" + encodeURIComponent(valueOf("hPartyCode")) + "&AgentEntryID=" + encodeURIComponent(valueOf("hAgentEntryID"));
			frm.submit();
			return true;
		};
		installSimpleUnload("Done");
	}

	function payRecRoot() {
		var root = xmlRoot("AccHeadData");
		return firstElement(root, "PayRec") || firstElement(root);
	}

	function updatePayRecCount(root) {
		var recCount = firstElement(xmlRoot("AccHeadData"), "RecCount") || root && root.selectSingleNode && root.selectSingleNode("//RecCount");
		if (recCount) {
			setAttribute(recCount, "Val", valueOf("hRecCount", "", "frm1"));
		}
	}

	function installPayRecSelection() {
		window.finaldone = function () {
			var root = xmlRoot("AccHeadData");
			var payRec = payRecRoot();
			var selectedCount = 0;
			if (!root || !payRec) {
				returnAndClose(root);
				return;
			}
			Array.prototype.forEach.call(form("frm1").elements, function (element) {
				var docNo;
				var parts;
				var docNode;
				var docValue;
				if (element.type === "checkbox" && element.name.indexOf("chkDocument") >= 0 && element.checked) {
					selectedCount = 1;
					docNo = element.value;
					parts = String(valueOf("hDoc" + docNo, "", "frm1")).split("?");
					docValue = String(docNo).split("Z");
					docNode = createXmlElement("AccHeadData", "Doc");
					setAttribute(docNode, "No", trim(docValue[0]));
					setAttribute(docNode, "InvNo", trim(parts[0]));
					setAttribute(docNode, "InvDate", trim(parts[1]));
					setAttribute(docNode, "TransAmount", formatNumber(parts[2]));
					setAttribute(docNode, "AmtAdjusted", formatNumber(parts[3]));
					setAttribute(docNode, "AmtToAdjust", "0");
					setAttribute(docNode, "DocType", valueOf("hVouType", "", "frm1"));
					setAttribute(docNode, "AmtToAccount", formatNumber(parts[4]));
					setAttribute(docNode, "PayableNo", trim(parts[5]));
					setAttribute(docNode, "AdjType", trim(parts[6]));
					payRec.appendChild(docNode);
				}
			});
			setAttribute(root, "No", selectedCount);
			updatePayRecCount(payRec);
			returnAndClose(root);
		};
		window.finalcancel = function () {
			var root = xmlRoot("AccHeadData");
			updatePayRecCount(payRecRoot());
			returnAndClose(root);
		};
		window.window_onunload = function () {
			if (!sent) {
				updatePayRecCount(payRecRoot());
				returnValue(xmlRoot("AccHeadData"));
			}
		};
		installEscape(window.finalcancel);
	}

	function tableById(id) {
		return document.getElementById(id) || document.all && document.all[id] || null;
	}

	function addCell(row, className, align, content) {
		var cell = row.insertCell(-1);
		if (className) {
			cell.className = className;
		}
		if (align) {
			cell.align = align;
		}
		if (content != null) {
			if (typeof content === "string") {
				cell.innerHTML = content;
			} else {
				cell.appendChild(content);
			}
		}
		return cell;
	}

	function installAgentCommission() {
		function root() {
			return xmlRoot("OutData");
		}
		window.ClearTable = function () {
			var tbl = tableById("tblBin") || tableById("tblbin");
			var row;
			if (!tbl) {
				return;
			}
			while (tbl.rows.length) {
				tbl.deleteRow(0);
			}
			row = tbl.insertRow(0);
			addCell(row, "ExcelHeaderCell", "center", "S.No.").width = "10";
			addCell(row, "ExcelHeaderCell", "center", "Agent Name");
			addCell(row, "ExcelHeaderCell", "center", "Commision type");
			addCell(row, "ExcelHeaderCell", "center", "Commision");
		};
		window.clearXML = function () {
			clearChildren(root());
		};
		window.AddShow = window.Addshow = function () {
			var tbl = tableById("tblBin") || tableById("tblbin");
			var target = field("selTobox") || field("selToBox");
			var xmlRootNode = root();
			if (!tbl || !target || !xmlRootNode) {
				return;
			}
			Array.prototype.forEach.call(target.options, function (option, index) {
				var row = tbl.insertRow(tbl.rows.length);
				var agentName = option.text;
				var textInput = document.createElement("input");
				var select = document.createElement("select");
				var commInput = document.createElement("input");
				var node = createXmlElement("OutData", "Agent");
				textInput.type = "text";
				textInput.value = agentName;
				textInput.name = "txtAgentname" + (index + 1);
				textInput.className = "FormElemRead";
				textInput.readOnly = true;
				[
					["Rate Per Quantity", "Q"],
					["Percentage on BV", "B"],
					["Percentage on IV", "V"]
				].forEach(function (entry) {
					select.options[select.options.length] = new Option(entry[0], entry[1]);
				});
				select.name = "cmbComtype" + (index + 1);
				select.className = "FormElem";
				commInput.type = "text";
				commInput.value = "";
				commInput.name = "txtComm" + (index + 1);
				commInput.className = "FormElem";
				addCell(row, "ExcelSerial", "center", String(index + 1));
				addCell(row, "ExcelDisplayCell", "left", textInput);
				addCell(row, "ExcelFieldCell", "center", select).width = "10";
				addCell(row, "ExcelInputCell", "center", commInput).width = "10";
				setAttribute(node, "Agentcode", option.value);
				setAttribute(node, "Agentname", agentName);
				setAttribute(node, "Commisiontype", "");
				setAttribute(node, "Commision", "");
				setAttribute(node, "CommValue", "");
				setAttribute(node, "PartyType", valueOf("hParType"));
				setAttribute(node, "PartySubType", valueOf("hParSubType"));
				xmlRootNode.appendChild(node);
			});
		};
		window.Newadd = function () {
			if (window.addclick) {
				window.addclick("selTobox", "selFrombox", "remove");
			}
			window.clearXML();
			window.ClearTable();
			window.AddShow();
		};
		window.Removeadd = function () {
			if (window.removeclick) {
				window.removeclick("selTobox", "selFrombox", "remove");
			}
			window.clearXML();
			window.ClearTable();
			window.AddShow();
		};
		window.Finalsubmit = window.FinalSubmit = function () {
			var nodes = childElements(root());
			var valid = true;
			nodes.forEach(function (node, index) {
				var type = field("cmbComtype" + (index + 1));
				var comm = field("txtComm" + (index + 1));
				if (!valid) {
					return;
				}
				if (!type || type.value === "0") {
					alert("Select Commision Type ");
					if (type) {
						type.focus();
					}
					valid = false;
					return;
				}
				if (!comm || trim(comm.value) === "") {
					alert("Enter Commission");
					if (comm) {
						comm.focus();
					}
					valid = false;
					return;
				}
				if (isNaN(Number(comm.value))) {
					alert("Enter Numbers Only");
					comm.select();
					valid = false;
					return;
				}
				setAttribute(node, "Commisiontype", type.value);
				setAttribute(node, "Commision", comm.value);
			});
			if (valid) {
				returnAndClose(root());
			}
		};
		window.FinalCancel = function () {
			window.clearXML();
			returnAndClose(root());
		};
		window.window_onunload = function () {
			if (!sent) {
				returnValue(root());
			}
		};
		installEscape(window.FinalCancel);
	}

	function installPartySubtypeDialog() {
		function sourceRoot() {
			return xmlRoot(ensureDialogArgs());
		}
		window.DisplaySubType = function () {
			var root = sourceRoot();
			var crTable = tableById("CRTable");
			var drTable = tableById("DRTable");
			var crCount = 0;
			var drCount = 0;
			if (!root || !crTable || !drTable) {
				return;
			}
			childElements(root, "Party").forEach(function (node) {
				var subType = String(getAttribute(node, "SubType")).split("|");
				var partyType = subType[0] || "";
				var partySubType = subType[1] || "";
				var row = (partyType === "CR" ? crTable : drTable).insertRow(-1);
				var check = getAttribute(node, "Check");
				var left = document.createElement("input");
				var right = document.createElement("input");
				left.type = "checkbox";
				left.name = (partyType === "CR" ? "ChkCR" + (++crCount) : "ChkDR" + (++drCount));
				left.value = partySubType + ":" + partyType;
				left.checked = check === "LCheck";
				right.type = "checkbox";
				right.name = partyType === "CR" ? "ChkPartySubType" : "ChkPartySubTypeDR";
				right.checked = check === "RCheck";
				addCell(row, "ExcelDisplayCell", "center", left);
				addCell(row, "ExcelDisplayCell", "Left", trim(node.textContent));
				addCell(row, "ExcelDisplayCell", "center", right);
			});
		};
		window.Init = function () {
			window.DisplaySubType();
		};
		window.CheckSubmit = window.window_onunload = function () {
			returnAndClose(sourceRoot());
		};
	}

	function installNewContact() {
		function textHost(id) {
			return document.getElementById(id) || window[id] || null;
		}

		function setText(id, value) {
			var host = textHost(id);
			if (host) {
				host.textContent = value == null ? "" : String(value);
			}
		}

		function rootFrom(value) {
			if (!value) {
				return null;
			}
			if (value.nodeType === 1) {
				return value;
			}
			return xmlRoot(value);
		}

		function serializeNode(node) {
			if (!node) {
				return "";
			}
			return new XMLSerializer().serializeToString(node.nodeType === 9 ? node : node.ownerDocument || node);
		}

		function contactRoot() {
			return xmlRoot("ContactData") || loadXmlIntoIsland("ContactData", "<Root/>");
		}

		function partyRoot() {
			return xmlRoot("PartyData") || loadXmlIntoIsland("PartyData", "<Party/>");
		}

		function setXmlChild(root, nodeName, value) {
			var node = firstElement(root, nodeName);
			if (!node) {
				node = createXmlElement("ContactData", nodeName);
				root.appendChild(node);
			}
			node.textContent = value == null ? "" : String(value);
			return node;
		}

		function checked(name) {
			var item = field(name);
			return !!(item && item.checked);
		}

		function parseResponseRoot(xhr) {
			if (xhr.responseXML && xhr.responseXML.documentElement) {
				return xhr.responseXML.documentElement;
			}
			if (trim(xhr.responseText)) {
				return new DOMParser().parseFromString(xhr.responseText, "text/xml").documentElement;
			}
			return null;
		}

		function attrValue(node, name) {
			return getAttribute(node, name);
		}

		function saveContactXml() {
			var root = contactRoot();
			var activeNode;
			var response;
			if (!root) {
				alert("Contact XML data is not available.");
				return false;
			}
			clearChildren(root);
			setXmlChild(root, "ParCode", valueOf("hContactNumber"));
			setXmlChild(root, "ParName", valueOf("txtContactName"));
			setXmlChild(root, "Designation", valueOf("txtDesignation"));
			setXmlChild(root, "ContactPersonFor", valueOf("txtContactPersonFor"));
			setXmlChild(root, "PartyCode", valueOf("hParCode"));
			setXmlChild(root, "Address1", valueOf("txtAddress1"));
			setXmlChild(root, "Address2", valueOf("txtAddress2"));
			setXmlChild(root, "PinCode", valueOf("txtPinCode"));
			setXmlChild(root, "City", valueOf("txtCity"));
			setXmlChild(root, "State", valueOf("txtState"));
			setXmlChild(root, "Country", valueOf("txtCountry"));
			setXmlChild(root, "Phone", valueOf("txtPhone"));
			setXmlChild(root, "Mobile", valueOf("txtMobileNo"));
			setXmlChild(root, "Fax", valueOf("txtFax"));
			setXmlChild(root, "Email", valueOf("txtEmail"));
			setXmlChild(root, "Website", valueOf("txtWebsite"));
			setXmlChild(root, "CreatedBy", valueOf("hCreatedBy"));
			activeNode = createXmlElement("ContactData", "Active");
			activeNode.setAttribute("Flag", checked("chkActive") ? "1" : "0");
			root.appendChild(activeNode);
			response = saveXmlIsland("XMLSaveParty.asp?Name=Contact&Mod=Master", "ContactData");
			if (trim(response) !== "") {
				alert(response);
				return false;
			}
			return true;
		}

		function handleSelectedParty(root) {
			var codes = [];
			var names = [];
			var details;
			var response;
			if (!root || String(getAttribute(root, "Action")).toUpperCase() !== "DONE") {
				return false;
			}
			details = childElements(root, "PartyDetails");
			if (!details.length) {
				return false;
			}
			details.forEach(function (node) {
				codes.push(attrValue(node, "ParCode"));
				names.push(attrValue(node, "ParName"));
			});
			response = saveXmlIsland("XMLSave.asp?Name=PartyType&Mod=SAL", "PartyData");
			if (trim(response) !== "") {
				alert(response);
				return false;
			}
			setValue("hPartyName", names.join(","));
			setValue("hParCode", codes.join(","));
			setText("spParty", names.join(","));
			return true;
		}

		function continuePartySelection(value) {
			var root = rootFrom(value);
			var action = String(getAttribute(root, "Action")).toUpperCase();
			var query = trim(getAttribute(root, "PassQuery"));
			if (action === "PAGE" && query) {
				openModernDialog("../Transaction/PartySelMultipleSubType.asp?" + query, xmlObject("PartyData"), "dialogHeight:500px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No", continuePartySelection);
				return;
			}
			handleSelectedParty(root);
		}

		function clearContactFields() {
			[
				"txtContactName",
				"txtDesignation",
				"txtContactPersonFor",
				"hPartyName",
				"hParCode",
				"txtAddress1",
				"txtAddress2",
				"txtCity",
				"txtPinCode",
				"txtState",
				"txtCountry",
				"txtPhone",
				"txtMobileNo",
				"txtFax",
				"txtEmail",
				"txtWebsite"
			].forEach(function (name) {
				setValue(name, "");
			});
			setText("spParty", "");
		}

		window.DelParty = function () {
			setValue("hPartyName", "");
			setValue("hParCode", "");
			setText("spParty", "");
		};

		window.SelPartyPopup = function () {
			var root = partyRoot();
			setValue("hPartyName", "");
			setText("spParty", "");
			if (root) {
				setAttribute(root, "Action", "");
				setAttribute(root, "PassQuery", "");
			}
			openModernDialog("../Transaction/PartySelMultipleSubType.asp?orgID=" + encodeURIComponent(valueOf("hUnitId")) + "&Party=0&hSelectMode=S", xmlObject("PartyData"), "dialogHeight:500px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No", continuePartySelection);
		};

		window.ViewData = function () {
			var frm = form("formname");
			if (trim(valueOf("hContactNumber")) === "") {
				alert("Party Details Cannot view because Party is not available");
				return false;
			}
			if (frm) {
				frm.action = "ParDetailsView.asp?PartyCode=" + encodeURIComponent(valueOf("hContactNumber"));
				frm.submit();
			}
			return true;
		};

		window.GoToMain = function () {
			var frm = form("formname");
			if (frm) {
				frm.action = "ContactsList.asp";
				frm.submit();
			}
			return true;
		};

		window.SaveXMLFinal = saveContactXml;
		window.Cleartxt = clearContactFields;

		window.popPartyDet = function (contactNo) {
			var xhr;
			var root;
			var node;
			var partyName;
			if (trim(contactNo) === "") {
				return;
			}
			xhr = new XMLHttpRequest();
			xhr.open("GET", "XMLGetContactData.asp?ContactNo=" + encodeURIComponent(contactNo), false);
			xhr.send(null);
			root = parseResponseRoot(xhr);
			if (!root) {
				return;
			}
			loadXmlIntoIsland("OutData", serializeNode(root));
			node = childElements(root)[0];
			if (!node) {
				return;
			}
			partyName = attrValue(node, "PartyName");
			setValue("txtDesignation", attrValue(node, "Designation"));
			setValue("txtContactPersonFor", attrValue(node, "ContactPersonFor"));
			setValue("hParCode", attrValue(node, "PartyCode"));
			setValue("hPartyName", partyName);
			setText("spParty", partyName);
			setValue("txtAddress1", attrValue(node, "AddressLine1"));
			setValue("txtAddress2", attrValue(node, "AddressLine2"));
			setValue("txtCity", attrValue(node, "City"));
			setValue("txtState", attrValue(node, "State"));
			setValue("txtCountry", attrValue(node, "Country"));
			setValue("txtPhone", attrValue(node, "PhoneNos"));
			setValue("txtMobileNo", attrValue(node, "MobileNos"));
			setValue("txtFax", attrValue(node, "FaxNos"));
			setValue("txtEmail", attrValue(node, "Email"));
			setValue("txtWebsite", attrValue(node, "WebsiteURL"));
			setValue("txtPinCode", attrValue(node, "Pincode").replace(/ /g, ""));
			setValue("txtContactName", node.textContent || "");
			if (field("chkActive")) {
				field("chkActive").checked = attrValue(node, "Useable") === "1";
			}
		};

		window.CheckForm = function () {
			var checks = [
				{ name: "txtContactName", message: "Enter Contact Name" },
				{ name: "txtDesignation", message: "Enter Designation" },
				{ name: "txtCity", message: "Enter City " }
			];
			var valid = true;
			checks.forEach(function (item) {
				var input;
				if (!valid || trim(valueOf(item.name)) !== "") {
					return;
				}
				alert(item.message);
				input = field(item.name);
				if (input && input.focus) {
					input.focus();
				}
				valid = false;
			});
			return valid;
		};

		window.PageSubmit = function () {
			var frm = form("formname");
			if (!window.CheckForm()) {
				return false;
			}
			if (!saveContactXml()) {
				return false;
			}
			if (frm) {
				frm.action = trim(valueOf("hContactNumber")) !== "" ? "ContactInsert.asp?Action=Edit" : "ContactInsert.asp?Action=Create";
				setButtonDisabled("B2", true);
				setButtonDisabled("B3", true);
				frm.submit();
			}
			return true;
		};
	}

	function installContactsList() {
		function submitTo(url) {
			var frm = form("formname");
			if (!frm) {
				return false;
			}
			if (url) {
				frm.action = url;
			}
			frm.submit();
			return true;
		}

		window.ViewContactDeatils = function (contactNo) {
			return submitTo("NewContact.asp?ContactNo=" + encodeURIComponent(contactNo));
		};
		window.ViewContactDetails = window.ViewContactDeatils;

		window.CreateNewContacts = function () {
			return submitTo("NewContact.asp");
		};

		window.AssignPage = function (pageNo) {
			setValue("hPage", pageNo);
			return submitTo("ContactsList.asp");
		};

		window.CheckSubmit = function () {
			var searchType = field("selParSearchType");
			setValue("hParName", valueOf("txtContactName"));
			setValue("hCity", valueOf("txtCity"));
			setValue("hSearch", searchType ? selectedValue(searchType) : "");
			return submitTo("ContactsList.asp");
		};

		window.DelContact = function () {
			var selected = fields("radButton").filter(function (radio) {
				return radio.checked;
			})[0];
			var response;
			if (!selected) {
				alert("Select the Contact to Delete");
				return false;
			}
			response = postText("ContactDeleteEntry.asp?hContactNo=" + encodeURIComponent(selected.value));
			if (trim(response) !== "") {
				alert(response);
			} else {
				alert("Contact Deleted Successfully");
				submitTo("ContactsList.asp");
			}
			return true;
		};
	}

	function init() {
		ensureDialogArgs();
		if (config.type === "unitSelection") {
			installUnitSelection();
		} else if (config.type === "partyDeleteSelection") {
			installPartyDeleteSelection();
		} else if (config.type === "partyPreference") {
			installPreferencePopup();
		} else if (config.type === "agentSelect") {
			installAgentSelect();
		} else if (config.type === "appBookUsed") {
			installAppBookUsed();
		} else if (config.type === "voucherTypeSelection") {
			installVoucherTypeSelection();
		} else if (config.type === "scheduleSelection") {
			installScheduleSelection();
		} else if (config.type === "scheduleSetupCaller") {
			installScheduleSetupCaller();
		} else if (config.type === "balanceScheduleSetup") {
			installBalanceScheduleSetup();
		} else if (config.type === "scheduleSubHeadsPopup") {
			installScheduleSubHeadsPopup();
		} else if (config.type === "plBsScheduleSubHeadsPopup") {
			installPlBsScheduleSubHeadsPopup();
		} else if (config.type === "scheduleBreakupSubHeadsPopup") {
			installScheduleBreakupSubHeadsPopup();
		} else if (config.type === "tdsComputation") {
			installTdsComputation();
		} else if (config.type === "tdsGroupingSetup") {
			installTdsGroupingSetup();
		} else if (config.type === "bankBookDetailsPopup") {
			installBankBookDetailsPopup();
		} else if (config.type === "dayBookGrid") {
			installDayBookGrid();
		} else if (config.type === "bookNarrations") {
			installBookNarrations();
		} else if (config.type === "narrationEntryPopup") {
			installNarrationEntryPopup();
		} else if (config.type === "booksEditEntryPopup") {
			installBooksEditEntryPopup();
		} else if (config.type === "booksCreationEntry") {
			installBooksCreationEntry();
		} else if (config.type === "glAccountHeadDetails") {
			installGlAccountHeadDetails();
		} else if (config.type === "glHeadCostCenterPopup") {
			installGlHeadScopedSelectionPopup("cost");
		} else if (config.type === "glHeadAnalyticalPopup") {
			installGlHeadScopedSelectionPopup("analytical");
		} else if (config.type === "glHeadPartySubTypePopup") {
			installGlHeadPartySubTypePopup();
		} else if (config.type === "glAccountHeadGrid") {
			installGlAccountHeadGrid();
		} else if (config.type === "partySubtypeEntry") {
			installPartySubtypeEntry();
		} else if (config.type === "autoClose") {
			installAutoClose();
		} else if (config.type === "miscPartyCreate") {
			installMiscPartyCreate();
		} else if (config.type === "representativeSelection") {
			installRepresentativeSelection();
		} else if (config.type === "representativeCreation") {
			installRepresentativeCreation();
		} else if (config.type === "contraList") {
			installContraList();
		} else if (config.type === "contraConfiguration") {
			installContraConfiguration();
		} else if (config.type === "partyContactPopup") {
			installPartyContactPopup();
		} else if (config.type === "partyLocationPopup") {
			installPartyLocationPopup();
		} else if (config.type === "partyCreateEditModals") {
			installPartyCreateEditModals();
		} else if (config.type === "partyUnitPopup") {
			installPartyUnitPopup();
		} else if (config.type === "payRecSelection") {
			installPayRecSelection();
		} else if (config.type === "agentCommission") {
			installAgentCommission();
		} else if (config.type === "partySubtypeDialog") {
			installPartySubtypeDialog();
		} else if (config.type === "newContact") {
			installNewContact();
		} else if (config.type === "contactsList") {
			installContactsList();
		}
	}

	init();
	window.ITMSPopupModernCompat = {
		init: init,
		returnValue: returnValue,
		returnAndClose: returnAndClose,
		xmlRoot: xmlRoot
	};
}(window, document));
