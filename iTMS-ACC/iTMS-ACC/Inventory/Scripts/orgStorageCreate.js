function trimTrue(val){
	var ltrim = /^\s+/g;
	var rtrim = /\s+$/g;
	return val.replace(ltrim,'').replace(rtrim,'');
}

function checkNumbers(val){
	var valid = "0123456789."
	var temp;

	for (var i=0; i < val.length; i++) {
		temp = "" + val.substring(i, i+1);
		if (valid.indexOf(temp) == "-1") 
			return false;
		else
			return true;
	}
}

function checkSubmit(){
	if (trimTrue(document.forms[0].txtLocationName.value) == "") {
		alert("Enter Location Name");
		document.forms[0].txtLocationName.select();
		return false;
	}
	else if (trimTrue(document.forms[0].txtLocationCode.value) == "") {
		alert("Select Location Code");
		document.forms[0].txtLocationCode.select();
		return false;
	}
	else if (!(document.forms[0].App[0].checked || (document.forms[0].App[1].checked) || (document.forms[0].App[2].checked) || (document.forms[0].App[3].checked) || (document.forms[0].App[4].checked) || (document.forms[0].App[5].checked) || (document.forms[0].App[6].checked) || (document.forms[0].App[7].checked))) {
		alert("Select Applicable For");
		return false;
	}
	else if (!(document.forms[0].ST[0].checked || (document.forms[0].ST[1].checked))) {
		alert("Select Storage Type");
		return false;
	}
	else if (document.forms[0].ST[0].checked && (trimTrue(document.forms[0].txtUsable.value) == "")) {
		alert("Enter Usable Free Area");			
		document.forms[0].txtUsable.select();
		return false;
	}

	else if (!(trimTrue(document.forms[0].txtUsable.value) == "") && (!checkNumbers(document.forms[0].txtUsable.value))) {
		alert("Enter Numerals Only");
		document.forms[0].txtUsable.select();
		return false;
	}
	else if (document.forms[0].ST[1].checked && (trimTrue(document.forms[0].txtBins.value) == "")) {
		alert("Enter Number of Bins");			
		document.forms[0].txtBins.select();
		return false;
	}

	else if (!(trimTrue(document.forms[0].txtBins.value) == "") && (!checkNumbers(document.forms[0].txtBins.value))) {
		alert("Enter Numerals Only");
		document.forms[0].txtBins.select();
		return false;
	}
	else {
		document.forms[0].action = "OrgStorageDefinitionInsert.asp"
		document.forms[0].submit();
	}
}

(function (window, document) {
	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function field(name) {
		var frm = form();
		return document.getElementById(name) || (frm && frm.elements ? frm.elements[name] : null);
	}

	function valueOf(name) {
		var item = field(name);
		return item ? item.value : "";
	}

	function xmlIsland(name) {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
		return window[name] || document[name] || document.getElementById(name);
	}

	function xmlDoc(name) {
		var item = xmlIsland(name);
		return item && item.XMLDocument || item && item._doc || null;
	}

	function xmlRoot(name) {
		var item = xmlIsland(name);
		return item && item.documentElement || item && item.XMLDocument && item.XMLDocument.documentElement || item && item._doc && item._doc.documentElement || null;
	}

	function serializeXml(nodeOrDoc) {
		var target = nodeOrDoc && nodeOrDoc.nodeType === 9 ? nodeOrDoc : nodeOrDoc && nodeOrDoc.ownerDocument || nodeOrDoc;
		return target ? new XMLSerializer().serializeToString(target) : "";
	}

	function saveNewData() {
		var doc = xmlDoc("NewData");
		var xhr = new XMLHttpRequest();
		xhr.open("POST", "XMLSave.asp?SessionFlag=True&Value=StorageNew&Folder=Master", false);
		xhr.send(serializeXml(doc));
	}

	function selectNodes(node, expression) {
		return node && node.selectNodes ? node.selectNodes(expression) : [];
	}

	function groupItem(name, index) {
		var group = field(name);
		return group && group[index] || null;
	}

	function selectedGroupValue(name, max) {
		var item;
		for (var i = 0; i < max; i += 1) {
			item = groupItem(name, i);
			if (item && item.checked) {
				return item.value;
			}
		}
		return "";
	}

	function setGroupByValue(name, max, value, disabled) {
		var item;
		for (var i = 0; i < max; i += 1) {
			item = groupItem(name, i);
			if (!item) {
				continue;
			}
			item.checked = item.value === value;
			if (disabled !== undefined) {
				item.disabled = !!disabled;
			}
		}
	}

	function removeTemporaryStorage(root) {
		var storages = selectNodes(root, "//Organization/Storage");
		var node;
		for (var i = storages.length - 1; i >= 0; i -= 1) {
			node = storages[i];
			if (trim(node.getAttribute("LOCATIONNUMBER")) === "0" && node.parentNode) {
				node.parentNode.removeChild(node);
			}
		}
	}

	function firstStorage() {
		var root = xmlRoot("NewData");
		var nodes = selectNodes(root, "//Organization/Storage");
		return nodes.length ? nodes[0] : null;
	}

	function FnInit() {
		var root = xmlRoot("NewData");
		var organizations;
		var storage;
		var flag = valueOf("hFlag") === "True";
		var img = field("Img2");
		if (!root) {
			return false;
		}
		organizations = selectNodes(root, "//Organization");
		for (var i = 0; i < organizations.length; i += 1) {
			organizations[i].setAttribute("UNITNAME", valueOf("hOrgName"));
		}
		storage = firstStorage();
		if (storage) {
			field("txtLocationName").value = storage.getAttribute("LOCATIONNAME") || "";
			field("txtLocationCode").value = storage.getAttribute("LOCATIONCODE") || "";
			setGroupByValue("App", 8, storage.getAttribute("APPLICABLEFOR") || "", flag);
			if (trim(storage.getAttribute("STORAGETYPEFREE")) === "0") {
				groupItem("ST", 1).checked = true;
				field("txtBins").value = storage.getAttribute("NUMBEROFBINS") || "";
				field("txtBins").disabled = false;
				if (img) {
					img.disabled = false;
				}
			} else if (trim(storage.getAttribute("STORAGETYPEBINS")) === "0") {
				groupItem("ST", 0).checked = true;
				field("txtUsable").value = storage.getAttribute("USABLEFREEAREA") || "";
			}
			if (flag) {
				groupItem("ST", 0).disabled = true;
				groupItem("ST", 1).disabled = true;
				field("txtUsable").disabled = true;
				field("txtBins").disabled = true;
				if (img) {
					img.disabled = true;
				}
			}
		}
		saveNewData();
		return false;
	}

	function openDetails() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog("XMLorgStoreView.asp", "A", "dialogHeight:510px;dialogWidth:620px;center:Yes;help:No;resizable:No;status:No");
		}
		return false;
	}

	function SetBinEnable(value) {
		var img = field("Img2");
		if (trim(value) === "B") {
			field("txtBins").disabled = false;
			field("txtUsable").value = "";
			if (img) {
				img.disabled = false;
			}
		} else {
			field("txtBins").disabled = true;
			field("txtBins").value = "";
			if (img) {
				img.disabled = true;
			}
		}
		return false;
	}

	function buildStorageElement(doc, locationNumber) {
		var storage = doc.createElement("Storage");
		var isFree = groupItem("ST", 0) && groupItem("ST", 0).checked;
		storage.setAttribute("LOCATIONNUMBER", locationNumber || "0");
		storage.setAttribute("LOCATIONCODE", valueOf("txtLocationCode"));
		storage.setAttribute("LOCATIONNAME", valueOf("txtLocationName"));
		storage.setAttribute("APPLICABLEFOR", selectedGroupValue("App", 8));
		storage.setAttribute("STORAGETYPEFREE", isFree ? "1" : "0");
		storage.setAttribute("STORAGETYPEBINS", isFree ? "0" : "1");
		storage.setAttribute("USABLEFREEAREA", valueOf("txtUsable"));
		storage.setAttribute("NUMBEROFBINS", valueOf("txtBins"));
		return storage;
	}

	function ensureOrganization(root, doc) {
		var org = selectNodes(root, "//Organization")[0];
		if (!org) {
			org = doc.createElement("Organization");
			root.appendChild(org);
		}
		org.setAttribute("OUDEFINITIONID", valueOf("hOrgID"));
		org.setAttribute("UNITNAME", valueOf("hOrgName"));
		return org;
	}

	function popBinSelect() {
		var root = xmlRoot("NewData");
		var doc = xmlDoc("NewData");
		var org;
		var storage;
		if (trim(valueOf("txtBins")) === "") {
			alert("Enter No. Of Bins");
			return false;
		}
		if (!root || !doc) {
			return false;
		}
		removeTemporaryStorage(root);
		if (trim(valueOf("hPara")) === "") {
			org = ensureOrganization(root, doc);
			org.appendChild(buildStorageElement(doc, "0"));
		} else {
			storage = firstStorage();
			if (storage) {
				storage.setAttribute("NUMBEROFBINS", valueOf("txtBins"));
			}
		}
		saveNewData();
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog("OrgStorageBinDetailsEntry.asp", xmlIsland("NewData"), "dialogHeight:500px;dialogWidth:450px;center:Yes;status:No", function () {
				var updatedStorage = firstStorage();
				if (updatedStorage) {
					field("txtBins").value = updatedStorage.getAttribute("NUMBEROFBINS") || "";
				}
			});
		}
		return false;
	}

	function clearChildren(node) {
		while (node && node.firstChild) {
			node.removeChild(node.firstChild);
		}
	}

	function XmlUpDate() {
		var root = xmlRoot("NewData");
		var doc = xmlDoc("NewData");
		var org;
		if (!root || !doc) {
			return false;
		}
		if (trim(valueOf("txtBins")) === "") {
			clearChildren(root);
			org = ensureOrganization(root, doc);
			org.appendChild(buildStorageElement(doc, ""));
		}
		saveNewData();
		return false;
	}

	window.FnInit = FnInit;
	window.openDetails = openDetails;
	window.SetBinEnable = SetBinEnable;
	window.popBinSelect = popBinSelect;
	window.XmlUpDate = XmlUpDate;
}(window, document));
