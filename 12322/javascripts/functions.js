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

var GoGlobal = new function() {
	self = this;
	var $proxy = "ajax/goglobalproxy.asp";
	var $siteGuid = 0;
	var $loader;
	var $message;

	self.countryList = [];

	self.Init = function() {
		$loader = $('.loader');
		$message = $('.message');

		self.getSiteList(self.handleResponse);
		self.getCountryList();
	}

	self.setSiteGuid = function(val) {
		$siteGuid = val;
	}

	self.getCurrentSiteGuid = function() {
		return $siteGuid;
	}

	self.getSiteList = function(callback) {
		$loader.show();
		$message.hide();

		//var $this = this;
		var result = null;
		
		$.post($proxy, {f:'loadSites'}, function(data, status, xhr) {
			result = $.parseJSON(data);

			xSites = "";
			result.forEach(function(aSite) {
				xSites += "<option value=" + aSite.Guid + ">" + aSite.Name + "</option>";
			});

			$('#ddlSiteList').append(xSites);

			$loader.hide();
			$message.hide();

			if(typeof callback === "function") {
				callback(xhr, data);
			}

		}).fail(function(data) {
			showResponseMessage(data);
		});

		return result;
	}

	self.getCountryList = function() {
		$loader.show();
		$message.hide();

		var obj = null;
		
		$.post($proxy, {f:'loadCountry'}, function(data, status, xhr) {
			obj = $.parseJSON(data);

			self.countryList = obj;

			var countryList = "";
			obj.forEach(function(aCountry) {
				countryList += "<div class=\"gridRow\">" + 
									"<div class=\"gridCell cName\">" + aCountry.Name + "</div>" +
									"<div class=\"gridCell\">" + aCountry.Guid + "</div>" +
									"<div class=\"gridCell\">" + aCountry.IsoCode + "</div>" +
									"<div class=\"gridCell\">" + aCountry.IsEuMember + "</div>" +
								"</div>";
			});
			$('#tblCountryList').append(countryList);

			$loader.hide();
			$message.hide();

		}).fail(function(data) {
			showResponseMessage(data);
		});
	}

	self.getLangAndCurrency = function(param) {
		$loader.show();
		$message.hide();

		$.post($proxy, param, function(data, status, xhr) {
			var obj = $.parseJSON(data);

			xLang = "";
			for (var i = 0; i < obj.Languages.length; i++) {
				var lang = obj.Languages[i];
				xLang += "<div class='x-name'>" + lang.Language + " (" + lang.Guid + ")</div><div class='x-prop'>" + (lang.Standard == "True" ? "Standard" : "") + "</div>";
			}
			$('.gLanguages').html(xLang);

			xCurr = "";
			for (var i = 0; i < obj.Currencies.length; i++) {
				var curr = obj.Currencies[i];
				xCurr += "<div class='x-name'>" + curr.Currency + " (" + curr.Guid + ")</div><div class='x-prop'>" + (curr.Standard == "True" ? "Standard" : "")+ "</div>";
			}
			$('.gCurrencies').html(xCurr);

			/* should change to Handleresponse ? */
			self.getAvailableLangAndCurrency({f:'loadLC', siteguid:$('#ddlSiteList').val()});

			$loader.hide();
			$message.hide();				

		}).fail(function(data) {
			showResponseMessage(getErrorMessage(data));
		});
	};

	self.getAvailableLangAndCurrency = function(param) {
		$('#formAddData select').each(function(index) {
			$(this).empty();
		});

		$.post($proxy, param, function(data, status, xhr) {
			var obj = $.parseJSON(data);

			var itemList = "";
			obj.Languages.forEach(function(aLang) {
				itemList += "<option value=" + aLang.Guid + ">" + aLang.Language + "</option>";
			});
			$('#sLanguage').append(itemList);

			itemList = "";
			obj.Currencies.forEach(function(aCurr) {
				itemList += "<option value=" + aCurr.Guid + ">" + aCurr.Currency + "</option>";
			})
			$('#sCurrency').append(itemList);

		}).fail(function(data) {
			showResponseMessage(getErrorMessage(data));
		})
	}

	self.AddEntity = function(entityType) {
		var $select;
		var $params;

		var $entityType = entityType.toLowerCase();
		
		switch(entityType.toLowerCase()) {
			case "language":
				$select = $('#sLanguage').val();
				break;
			case "currency":
				$select = $('#sCurrency').val();				
				break;
			default:
				break;
		}

		$params = {f:'addEntity', entityType:$entityType, guid:$select, siteguid:$siteGuid};

		console.log($params);

		$loader.show();
		$message.hide();

		$.post($proxy, $params, function(data, status, xhr) {

			var xData = self.handleResponse(xhr, data);

			$loader.hide();

			if(xData.Error !== 1) {
				$message.hide();
			}

		}).fail(function(data) {
			showResponseMessage(getErrorMessage(data));
		})
	}

	self.processSiteSelection = function() {
		if (Location.Hash("siteguid") !== "" && $.isNumeric(Location.Hash("siteguid"))) {
			var selectedSiteGuid = Location.Hash("siteguid");
			
			//Load data based on current query string
			//self.getLangAndCurrency({f:'loadData', siteguid:selectedSiteGuid});
			$('#ddlSiteList').val(selectedSiteGuid).change();
		}
	}

	self.handleResponse = function(xhr, data) {
		var headerJSON = getJSON(xhr);
		//var retObj = null;

		if (typeof(headerJSON) == 'object' && headerJSON) {
			switch(headerJSON.EntityType) {
				case "Site":
					self.processSiteSelection();
					break;
				case "UpdateEntityTable":
					self.getLangAndCurrency({f:'loadData', siteguid:$siteGuid})
					break;
				case "ShowMessage":
					showResponseMessage(headerJSON.MessageText);
					break;
				default:
					break;
			}
		}
		return $.parseJSON(data);
	}

	var showResponseMessage = function(response) {
		$loader.hide();
		$message.show().html(response.MessageText);
	}

	var getErrorMessage = function(data) {
		return {"MessageText":"Error code: " + data.status + ".<br/>Message: " + data.statusText};
	}
};


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

$(document).ready(function() {
	
	// Form initialize
	GoGlobal.Init();
	
	/* Event binding */
	$('#ddlSiteList').on('change', function() {
		GoGlobal.setSiteGuid($(this).val());

		GoGlobal.getLangAndCurrency($('#formGoGlobalView').eq(0).serialize());

		//Update query string with new siteguid parameter
		Location.UpdateURLHash([{"key":"siteguid", "value": $(this).val()}]);
	});

	$('#btnAddLanguage').on('click', function() {
		GoGlobal.AddEntity("language");
	})

	$('#btnAddCurrency').on('click', function() {
		GoGlobal.AddEntity("currency");	
	})

	/* e:Event binding */
});