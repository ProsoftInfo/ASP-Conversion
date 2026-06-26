function checkNull(val){
	var ltrim = /^\s+/;
	var rtrim = /\s+$/;
	val = val.replace(ltrim,'');
	val = val.replace(rtrim,'');
	return val;
}
function checkNumber(val){
	var num = /[0-9]+[.]{0,2}[0-9]*$/g
	var num1 = /^[0-9]+/
	var num2 = /[0-9]+$/
	var num3 = /[0-9]*[.]{1,}[0-9]*[.]{1,}[0-9]*/g

	if ( !num1.test(val) || !num2.test(val) || num3.test(val)){
		return false;
	}else return true;
}

