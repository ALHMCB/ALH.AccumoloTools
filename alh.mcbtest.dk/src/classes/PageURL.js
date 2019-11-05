let PageURL = class {
	constructor() {}

	static getHash(item) {
		const val = location.hash.match(new RegExp("[\#\&]" + item + "=([^\&]*)(\&?)","i"));
		return val ? val[1] : val;
	}

	static parseURLHash() {
		var data = {};
		if (window.location.hash !== "") {
			var params = window.location.hash.split("#")[1].split("&");
			params.forEach(function(pair) {
				pair = pair.split("=");
				data[pair[0]] = decodeURIComponent(pair[1]||'');
			});
		}
		return data;
	}

	static updateURLHash(params) {
		var xJSON = this.parseURLHash();
		params.forEach(function(oParam) {
			xJSON[oParam.key] = oParam.value;
		});
		var newSearch = "#" + $.param(xJSON);
		window.location.hash = newSearch;
	}
}