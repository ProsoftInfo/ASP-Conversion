function trimValue(value) {
	return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
}

function getMasterForm() {
	return document.forms.formname || document.forms[0];
}

function selectedText(selectElement) {
	if (!selectElement || selectElement.selectedIndex < 0) {
		return "";
	}
	return selectElement.options[selectElement.selectedIndex].text;
}

function selectedParts(selectElement) {
	if (!selectElement) {
		return [];
	}
	return String(selectElement.value || "").split("|");
}

function setRadioValue(radios, value) {
	if (!radios) {
		return;
	}
	if (radios.length == null) {
		radios.checked = radios.value === value;
		return;
	}
	for (var index = 0; index < radios.length; index += 1) {
		radios[index].checked = radios[index].value === value;
	}
}

function hasCategoryFields(form) {
	return !!(form && form.selCategory);
}

function hasUomFields(form) {
	return !!(form && form.selUoM);
}

function checkCategorySubmit(form) {
	var isNew = form.selCategory.selectedIndex === 0;
	form.hFlag.value = isNew ? "N" : "A";

	if (isNew && trimValue(form.txtCatCode.value) === "") {
		alert("Enter Category Code");
		form.txtCatCode.select();
		return false;
	}
	if (trimValue(form.txtCatName.value) === "") {
		alert("Enter Category Name");
		form.txtCatName.select();
		return false;
	}
	if (trimValue(form.txtCatShName.value) === "") {
		alert("Enter Category Short Name");
		form.txtCatShName.select();
		return false;
	}

	form.action = "MasCategoryUpdate.asp";
	form.submit();
	return true;
}

function checkUomSubmit(form) {
	if (form.selUoM.selectedIndex === 0) {
		alert("Select UoM");
		form.selUoM.focus();
		return false;
	}
	if (trimValue(form.txtUOMName.value) === "") {
		alert("Enter UoM description");
		form.txtUOMName.select();
		return false;
	}
	if (!form.radDecimal || !form.radDecimal.value) {
		alert("Select Decimals Allowed");
		return false;
	}

	form.action = "masUOMUpdate.asp";
	form.submit();
	return true;
}

function checkSubmit() {
	var form = getMasterForm();
	if (hasCategoryFields(form)) {
		return checkCategorySubmit(form);
	}
	if (hasUomFields(form)) {
		return checkUomSubmit(form);
	}
	return false;
}

function populateCategoryDetails(form, selectElement) {
	if (selectElement.selectedIndex === 0) {
		form.hFlag.value = "N";
		form.txtCatCode.value = "";
		form.txtCatName.value = "";
		form.txtCatShName.value = "";
		return false;
	}

	var parts = selectedParts(selectElement);
	form.hFlag.value = "A";
	form.txtCatCode.value = trimValue(parts[0]);
	form.txtCatName.value = trimValue(selectedText(selectElement));
	form.txtCatShName.value = trimValue(parts[1]);
	return true;
}

function populateUomDetails(form, selectElement) {
	if (selectElement.selectedIndex === 0) {
		form.txtUOMCode.value = "";
		form.txtUOMName.value = "";
		setRadioValue(form.radDecimal, "");
		return false;
	}

	var parts = selectedParts(selectElement);
	form.txtUOMCode.value = trimValue(parts[1] || parts[0]);
	form.txtUOMName.value = trimValue(selectedText(selectElement));
	setRadioValue(form.radDecimal, trimValue(parts[2]));
	return true;
}

function GetDetails(obj) {
	var form = getMasterForm();
	if (hasCategoryFields(form)) {
		return populateCategoryDetails(form, obj);
	}
	if (hasUomFields(form)) {
		return populateUomDetails(form, obj);
	}
	return false;
}

function deleteCategory(form) {
	if (form.selCategory.selectedIndex === 0) {
		alert("Select Category");
		form.selCategory.focus();
		return false;
	}

	form.hCatName.value = trimValue(selectedText(form.selCategory));
	form.action = "CategoryDeletionUpdate.asp";
	form.submit();
	return true;
}

function deleteUom(form) {
	if (form.selUoM.selectedIndex === 0) {
		alert("Select UoM");
		form.selUoM.focus();
		return false;
	}

	form.hUoMName.value = trimValue(selectedText(form.selUoM));
	form.action = "UoMDeletionUpdate.asp";
	form.submit();
	return true;
}

function Delete() {
	var form = getMasterForm();
	if (hasCategoryFields(form)) {
		return deleteCategory(form);
	}
	if (hasUomFields(form)) {
		return deleteUom(form);
	}
	return false;
}
