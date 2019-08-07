var GoGlobal = new function() {
	let self = this;

	this.countryList = [];

	this.Init = function() {
		SiteUtils.setTargetDropdownList($('#ddlSiteList'));
		SiteUtils.Init();
		self.getCountryList();
	}

	this.getCountryList = function() {
		SiteUtils.showLoader();
		SiteUtils.hideMessage();

		var obj = null;
		
		$.post($proxy, {f:'loadCountry'}, function(data, status, xhr) {
			obj = $.parseJSON(data);

			SiteUtils.countryList = obj;

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

			SiteUtils.hideLoader();
			SiteUtils.hideMessage();

		}).fail(function(data) {
			SiteUtils.showResponseMessage(data);
		});
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
	
	/* Event binding */
	$('#ddlSiteList').on('change', function() {
		SiteUtils.setSiteGuid($(this).val());

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