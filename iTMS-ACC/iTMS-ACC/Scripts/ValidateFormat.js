function itmsRememberKeyEvent(evt) {
	window.__itmsCurrentKeyEvent = evt || null;
}

if (document.addEventListener) {
	document.addEventListener("keypress", itmsRememberKeyEvent, true);
	document.addEventListener("keydown", itmsRememberKeyEvent, true);
}

function itmsCurrentKeyEvent(evt) {
	return evt || window.__itmsCurrentKeyEvent || null;
}

function itmsEventTarget(evt) {
	return evt && (evt.target || evt.srcElement) || null;
}

function itmsKeyCode(evt) {
	return evt ? evt.which || evt.keyCode || evt.charCode || 0 : 0;
}

function itmsCancelKey(evt) {
	if (!evt) {
		return false;
	}
	if (evt.preventDefault) {
		evt.preventDefault();
	}
	evt.returnValue = false;
	return false;
}

function itmsAllowControlKey(code) {
	return code === 0 || code === 8 || code === 9 || code === 13 || code === 27;
}

function DoKeyPress(sYesNo, iIntPart, iDecPart, evt) {
	var eventObj = itmsCurrentKeyEvent(evt);
	var target = itmsEventTarget(eventObj);
	var code = itmsKeyCode(eventObj);
	var value = target && target.value != null ? String(target.value) : "";
	var decPosition = value.indexOf(".");
	var intValue = decPosition >= 0 ? value.substring(0, decPosition) : value;
	var decValue = decPosition >= 0 ? value.substring(decPosition + 1) : "";

	if (itmsAllowControlKey(code)) {
		return true;
	}
	if (sYesNo === "N" && (code < 48 || code > 57)) {
		return itmsCancelKey(eventObj);
	}
	if (sYesNo === "Y" && ((code < 48 || code > 57) && code !== 46)) {
		return itmsCancelKey(eventObj);
	}
	if (sYesNo === "N" && intValue.length >= iIntPart) {
		return itmsCancelKey(eventObj);
	}
	if (sYesNo === "Y") {
		if (decPosition >= 0) {
			if (code === 46 || decValue.length >= iDecPart) {
				return itmsCancelKey(eventObj);
			}
		} else if (intValue.length >= iIntPart && code !== 46) {
			return itmsCancelKey(eventObj);
		}
	}
	return true;
}

function DoKeyPressHypen(sYesNo, iIntPart, iDecPart, evt) {
	var eventObj = itmsCurrentKeyEvent(evt);
	var target = itmsEventTarget(eventObj);
	var code = itmsKeyCode(eventObj);
	var value = target && target.value != null ? String(target.value) : "";
	var decPosition = value.indexOf(".");
	var intValue = decPosition >= 0 ? value.substring(0, decPosition) : value;
	var decValue = decPosition >= 0 ? value.substring(decPosition + 1) : "";

	if (itmsAllowControlKey(code)) {
		return true;
	}
	if (sYesNo === "N" && ((code < 48 || code > 57) && code !== 45)) {
		return itmsCancelKey(eventObj);
	}
	if (sYesNo === "Y" && ((code < 48 || code > 57) && code !== 46 && code !== 45)) {
		return itmsCancelKey(eventObj);
	}
	if (sYesNo === "N" && intValue.length >= iIntPart) {
		return itmsCancelKey(eventObj);
	}
	if (sYesNo === "Y") {
		if (decPosition >= 0) {
			if (code === 46 || decValue.length >= iDecPart) {
				return itmsCancelKey(eventObj);
			}
		} else if (intValue.length >= iIntPart && code !== 46) {
			return itmsCancelKey(eventObj);
		}
	}
	return true;
}

function CheckAlpha(len, evt) {
	var eventObj = itmsCurrentKeyEvent(evt);
	var target = itmsEventTarget(eventObj);
	var code = itmsKeyCode(eventObj);
	var maxLength = Number(len) || 0;
	var value = target && target.value != null ? String(target.value) : "";

	if (itmsAllowControlKey(code)) {
		return true;
	}
	if (!((code >= 97 && code <= 122) || (code >= 65 && code <= 90))) {
		alert("Only Alphabets should be entered");
		return itmsCancelKey(eventObj);
	}
	if (maxLength > 0 && value.length >= maxLength) {
		return itmsCancelKey(eventObj);
	}
	return true;
}

function DoKeyPressText(iLength, evt) {
	var eventObj = itmsCurrentKeyEvent(evt);
	var target = itmsEventTarget(eventObj);
	var maxLength = Number(iLength) || 0;
	var value = target && target.value != null ? String(target.value) : "";

	if (target && maxLength > 0 && value.length > maxLength) {
		target.value = value.substring(0, maxLength);
	}
	return true;
}
