var Location = new function () {
	
	this.Hash = function(item) {
		var svalue = location.hash.match(new RegExp("[\#\&]" + item + "=([^\&]*)(\&?)","i"));
		return svalue ? svalue[1] : svalue;
	};

	this.ParseURLHash = function() {
		var data = {};
		if (window.location.hash !== "") {
			var params = window.location.hash.split("#")[1].split("&");
			params.forEach(function(pair) {
				pair = pair.split("=");
				data[pair[0]] = decodeURIComponent(pair[1]||'');
			});
		}
		return data;
	};

	this.UpdateURLHash = function(params) {
		var xJSON = this.ParseURLHash();
		params.forEach(function(oParam) {
			xJSON[oParam.key] = oParam.value;
		});
		var newSearch = "#" + $.param(xJSON);
		window.location.hash = newSearch;
	};
}

var SiteUtils = new function() {
	let self = this;
	var $siteGuid = 0;
	var $targetDropdown;
	var $loader;
	var $message;

	this.Init = function() {
		$loader = $('.loader');
		$message = $('.message');
		this.getSiteList(self.handleResponse);
	}

	this.setSiteGuid = function(val) {
		$siteGuid = val;
	}

	this.getCurrentSiteGuid = function() {
		return $siteGuid;
	}

	this.setTargetDropdownList = function(val) {
		$targetDropdown = val;
	}

	this.getSiteList = function(callback) {
		self.showLoader();
		self.hideMessage();

		var result = null;
		
		$.post($proxy, {f:'loadSites'}, function(data, status, xhr) {
			result = $.parseJSON(data);

			xSites = "";
			result.forEach(function(aSite) {
				xSites += "<option value=" + aSite.Guid + ">" + aSite.Name + "</option>";
			});

			$targetDropdown.append(xSites);

			self.hideLoader();
			self.hideMessage();

			if(typeof callback === "function") {
				callback(xhr, data);
			}

		}).fail(function(data) {
			self.showResponseMessage(data);
		});

		return result;
	}

	this.processSiteSelection = function() {
		if (Location.Hash("siteguid") !== "" && $.isNumeric(Location.Hash("siteguid"))) {
			var selectedSiteGuid = Location.Hash("siteguid");
			
			$targetDropdown.val(selectedSiteGuid).change();
		}
	}

	this.handleResponse = function(xhr, data) {
		var headerJSON = getJSON(xhr);

		if (typeof(headerJSON) == 'object' && headerJSON) {
			switch(headerJSON.EntityType) {
				case "Site":
					self.processSiteSelection();
					break;
				case "UpdateEntityTable":
					GoGlobal.getLangAndCurrency({f:'loadData', siteguid:$siteGuid})
					break;
				case "Domf":
					Domf.showDOMFStatus(headerJSON);
					break;
				case "UpdateDOMFStatus":
					Domf.getDOMFStatus({f:'loadDOMFStatus', siteguid:$siteGuid}, self.handleResponse);
					break;
				case "ShowMessage":
					self.showResponseMessage(headerJSON.MessageText);
					break;
				default:
					break;
			}
		}
		return $.parseJSON(data);
	}

	this.showLoader = function() {
		$loader.show();
	}

	this.hideLoader = function() {
		$loader.hide();
	}

	this.showMessage = function() {
		$message.show();
	}

	this.hideMessage = function() {
		$message.hide();
	}

	this.showResponseMessage = function(response) {
		this.hideLoader();
		$message.show().html(response.MessageText);
	}

	this.getErrorMessage = function(data) {
		return {"MessageText":"Error code: " + data.status + ".<br/>Status: " + data.statusText + ".<br/>Response: " + data.responseText};
	}
}


function getJSON(obj, headerName)
{
	if (typeof headerName === 'undefined') {
		headerName = 'X-JSON';
	}
	
	var strJson = obj.getResponseHeader(headerName);
	
	if (typeof obj.getResponseHeader(headerName) === 'string') {
		strJson = strJson.replace(/\\\"/g, '"');
	}
	
	return $.parseJSON(strJson);
}

