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

	function setValue(name, value) {
		var item = field(name);
		if (item) {
			item.value = value == null ? "" : String(value);
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

	function firstElement(name) {
		return childElements(xmlRoot(name))[0] || null;
	}

	function attr(node, name) {
		return node && node.getAttribute(name) || "";
	}

	function findByAttr(name, nodeName, attrName, attrValue) {
		var found = null;
		childElements(xmlRoot(name)).some(function (node) {
			if (node.nodeName === nodeName && attr(node, attrName) === attrValue) {
				found = node;
				return true;
			}
			return false;
		});
		return found;
	}

	function setSelectSingleOption(name, value, text) {
		var select = field(name);
		if (!select) {
			return;
		}
		select.options.length = 0;
		select.options[0] = new Option(text || "", value || "");
		select.selectedIndex = 0;
	}

	function setSelectIndexByValue(select, value) {
		if (!select) {
			return false;
		}
		for (var i = 0; i < select.options.length; i += 1) {
			if (trim(select.options[i].value) === trim(value)) {
				select.selectedIndex = i;
				return true;
			}
		}
		return false;
	}

	function setDateValue(value) {
		var picker = field("UnitCSTRCDate");
		if (picker) {
			if (typeof picker.SetDate === "function") {
				picker.SetDate(value);
			} else if (typeof picker.setDate === "function") {
				picker.setDate(value);
			} else {
				picker.value = value || "";
			}
		}
	}

	function setHiddenDate() {
		var picker = field("UnitCSTRCDate");
		if (!picker) {
			setValue("txtUnitCSTRCDate", "");
		} else if (typeof picker.GetDate === "function") {
			setValue("txtUnitCSTRCDate", picker.GetDate());
		} else if ("getDate" in picker && typeof picker.getDate === "function") {
			setValue("txtUnitCSTRCDate", picker.getDate());
		} else {
			setValue("txtUnitCSTRCDate", picker.value || "");
		}
	}

	function openDetails() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog("XMLUnitView.asp", "Unit", "dialogHeight:310px;dialogWidth:320px;center:Yes;help:No;resizable:No;status:No", function () {});
		} else {
			window.open("XMLUnitView.asp", "Unit", "width=320,height=310,resizable=no,scrollbars=yes");
		}
	}

	function fillCommonFields(node, replaceCountry) {
		setValue("txtUnitAddr1", attr(node, "ADDRESS1"));
		setValue("txtUnitAddr2", attr(node, "ADDRESS2"));
		setValue("txtUnitPIN", attr(node, "POSTCODE"));
		setValue("txtUnitCity", attr(node, "CITY"));
		setValue("txtUnitState", attr(node, "STATE"));
		if (replaceCountry) {
			setSelectSingleOption("selUnitCountry", attr(node, "COUNTRYCODE"), attr(node, "COUNTRYNAME"));
		} else {
			setSelectIndexByValue(field("selUnitCountry"), attr(node, "COUNTRYCODE"));
		}
		setValue("txtUnitPhone", attr(node, "PHONENUMBER"));
		setValue("txtUnitFax", attr(node, "FAXNUMBER"));
		setValue("txtUnitEmail", attr(node, "EMAILID"));
		setValue("txtUnitURL", attr(node, "WESITEURL"));
		setValue("txtUnitContactPerson", attr(node, "CONTACTPERSON"));
		setValue("txtUnitTNGSTNo", attr(node, "TNGSTRCNUMBER"));
		setValue("txtUnitCSTRCNo", attr(node, "CSTRCNUMBER"));
		setDateValue(attr(node, "CSTRCDATE"));
		setValue("txtUnitAreaCode", attr(node, "AREACODE"));
	}

	function fillExciseFields(node) {
		setValue("txtUnitRange", attr(node, "RANGE"));
		setValue("txtUnitDivision", attr(node, "DIVISION"));
		setValue("txtUnitCollectorate", attr(node, "COLLECTORATE"));
		setValue("txtUnitCentralENo", attr(node, "CENTRALEXCISECODE"));
		setValue("txtUnitRegNo", attr(node, "REGISTRATIONNUMBER"));
		setValue("txtUnitLANo", attr(node, "LANUMBER"));
		setValue("txtRangeAdd1", attr(node, "RANGEADDRESS1"));
		setValue("txtRangeAdd2", attr(node, "RANGEADDRESS2"));
		setValue("txtRangeAdd3", attr(node, "RANGEADDRESS3"));
		setValue("txtDivisionAdd1", attr(node, "DIVISIONADDRESS1"));
		setValue("txtDivisionAdd2", attr(node, "DIVISIONADDRESS2"));
		setValue("txtDivisionAdd3", attr(node, "DIVISIONADDRESS3"));
	}

	function displayOrgDet() {
		var frm = form();
		var orgUnit = field("selOrgUnit");
		var orgNode;
		if (!orgUnit) {
			return false;
		}
		if (orgUnit.selectedIndex === 1) {
			frm.reset();
			orgUnit.selectedIndex = 1;
			if (field("selParOrgUnit")) {
				field("selParOrgUnit").disabled = true;
			}
			orgNode = firstElement("orgData");
			fillCommonFields(orgNode, true);
		} else if (orgUnit.selectedIndex === 2) {
			frm.reset();
			orgUnit.selectedIndex = 2;
			if (field("selParOrgUnit")) {
				field("selParOrgUnit").disabled = false;
			}
		} else {
			frm.reset();
			if (field("selParOrgUnit")) {
				field("selParOrgUnit").disabled = false;
			}
		}
		return true;
	}

	function displayDivDet(select) {
		var node;
		if (field("selOrgUnit") && field("selOrgUnit").selectedIndex === 0) {
			alert("Select Organization Unit");
			field("selOrgUnit").focus();
			form().reset();
			return false;
		}
		node = findByAttr("DivData", "Division", "OUDEFINITIONID", select.options[select.selectedIndex].value);
		if (node) {
			fillCommonFields(node, true);
			fillExciseFields(node);
		}
		return true;
	}

	function loadUnits() {
		var org = field("selOrg");
		var orgUnit = field("selOrgUnit");
		var target = field("selDivisionUnit");
		var selectedOrg = org && org.selectedIndex;
		var selectedOrgUnit = orgUnit && orgUnit.selectedIndex;
		var sourceName;
		var nodeName;
		if (!target || !orgUnit) {
			return false;
		}
		form().reset();
		if (org) {
			org.selectedIndex = selectedOrg;
		}
		orgUnit.selectedIndex = selectedOrgUnit;
		sourceName = selectedOrgUnit === 1 ? "DivisionData" : selectedOrgUnit === 2 ? "UnitData" : "";
		nodeName = selectedOrgUnit === 1 ? "Division" : selectedOrgUnit === 2 ? "Unit" : "";
		target.length = 1;
		if (!sourceName) {
			return false;
		}
		childElements(xmlRoot(sourceName)).forEach(function (node) {
			target.options[target.options.length] = new Option(attr(node, "ORGUNITDESCRIPTION"), attr(node, "OUDEFINITIONID"));
		});
		return true;
	}

	function loadDetails() {
		var orgUnit = field("selOrgUnit");
		var selected = field("selDivisionUnit");
		var sourceName;
		var nodeName;
		var node;
		if (!orgUnit || !selected) {
			return false;
		}
		sourceName = orgUnit.selectedIndex === 1 ? "DivisionData" : orgUnit.selectedIndex === 2 ? "UnitData" : "";
		nodeName = orgUnit.selectedIndex === 1 ? "Division" : orgUnit.selectedIndex === 2 ? "Unit" : "";
		node = sourceName ? findByAttr(sourceName, nodeName, "OUDEFINITIONID", selected.value) : null;
		if (!node) {
			return false;
		}
		setValue("txtUnitName", attr(node, "ORGUNITDESCRIPTION"));
		setValue("txtUnitShName", attr(node, "ORGUNITSHORTDESCRIPTION"));
		fillCommonFields(node, false);
		fillExciseFields(node);
		return true;
	}

	function installEntry() {
		upgradeXml();
		window.openDetails = openDetails;
		window.DisplayOrgDet = displayOrgDet;
		window.DisplayDivDet = displayDivDet;
		window.sethiddenDate = setHiddenDate;
	}

	function installAmend() {
		upgradeXml();
		window.openDetails = openDetails;
		window.LoadUnits = loadUnits;
		window.LoadDetails = loadDetails;
		window.setIndex = setSelectIndexByValue;
		window.sethiddenDate = setHiddenDate;
	}

	window.ITMSOrgUnitDefinitionCompat = {
		installEntry: installEntry,
		installAmend: installAmend
	};
}(window, document));
