var GoGlobal = new function() {
	let self = this;

	this.countryList = [];

	this.Init = function() {
		SiteUtils.setTargetDropdownList($('#ddlSiteList'));
		SiteUtils.Init();
		self.getCountryList();

		/* Event binding */
		$('#ddlSiteList').on('change', function() {
			SiteUtils.setSiteGuid(this.value);

			self.getLangAndCurrency($('#formGoGlobalView').eq(0).serialize());

			//Update query string with new siteguid parameter
			Location.UpdateURLHash([{"key":"siteguid", "value": this.value}]);
		});

		$('#btnAddLanguage').on('click', function() {
			self.AddEntity("language");
		})

		$('#btnAddCurrency').on('click', function() {
			self.AddEntity("currency");	
		})

		$('#txtFilter').on('keyup', function() {
			self.printCountryList(this.value);
		})

		/* e:Event binding */
	}

	this.getCountryList = function() {
		SiteUtils.showLoader();
		SiteUtils.hideMessage();

		var obj = null;
		
		$.post($proxy, {f:'loadCountry'}, function(data, status, xhr) {
			obj = $.parseJSON(data);

			GoGlobal.countryList = obj;

			GoGlobal.printCountryList();

			SiteUtils.hideLoader();
			SiteUtils.hideMessage();

		}).fail(function(data) {
			SiteUtils.showResponseMessage(data);
		});
	}

	this.printCountryList = function(key = "") {
		const countryTable = $('#tblCountryList');
		countryTable.empty();

		if(GoGlobal.countryList != null && typeof(GoGlobal.countryList) == "object") {

			let countryList = "";

			if(key != "") {
				key = key.toLowerCase();
				
				GoGlobal.countryList.forEach(function(aCountry) {
					
					if(aCountry.Name.toLowerCase().indexOf(key) !== -1 || aCountry.Guid.indexOf(key) !== -1 || aCountry.IsoCode.toLowerCase().indexOf(key) !== -1) {
						countryList += "<div class=\"gridRow\">" + 
										"<div class=\"gridCell cName\">" + aCountry.Name + "</div>" +
										"<div class=\"gridCell\">" + aCountry.Guid + "</div>" +
										"<div class=\"gridCell\">" + aCountry.IsoCode + "</div>" +
										"<div class=\"gridCell\">" + aCountry.IsEuMember + "</div>" +
									"</div>";
					}
				});
			}
			else {
				GoGlobal.countryList.forEach(function(aCountry) {					
					countryList += "<div class=\"gridRow\">" + 
									"<div class=\"gridCell cName\">" + aCountry.Name + "</div>" +
									"<div class=\"gridCell\">" + aCountry.Guid + "</div>" +
									"<div class=\"gridCell\">" + aCountry.IsoCode + "</div>" +
									"<div class=\"gridCell\">" + aCountry.IsEuMember + "</div>" +
								"</div>";
				});
			}
				
			countryTable.append(countryList);
		}
	}

	this.getLangAndCurrency = function(param) {
		SiteUtils.showLoader();
		SiteUtils.hideMessage();

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

			SiteUtils.hideLoader();
			SiteUtils.hideMessage();

		}).fail(function(data) {
			SiteUtils.showResponseMessage(SiteUtils.getErrorMessage(data));
		});
	};

	this.getAvailableLangAndCurrency = function(param) {
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
			SiteUtils.showResponseMessage(SiteUtils.getErrorMessage(data));
		})
	}

	this.AddEntity = function(entityType) {
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

		$params = {f:'addEntity', entityType:$entityType, guid:$select, siteguid:SiteUtils.getCurrentSiteGuid()};

		console.log($params);

		SiteUtils.showLoader();
		SiteUtils.hideMessage();

		$.post($proxy, $params, function(data, status, xhr) {

			var xData = SiteUtils.handleResponse(xhr, data);

			SiteUtils.hideLoader();

			if(xData.Error !== 1) {
				SiteUtils.hideMessage();
			}

		}).fail(function(data) {
			SiteUtils.showResponseMessage(SiteUtils.getErrorMessage(data));
		})
	}
};

$(document).ready(function() {
	
	// Form initialize
	GoGlobal.Init();
});