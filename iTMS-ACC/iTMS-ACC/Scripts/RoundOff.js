function RndOff(iValue) {
	var value = Number(iValue);
	if (isNaN(value)) {
		return 0;
	}
	return Math.round(value);
}
