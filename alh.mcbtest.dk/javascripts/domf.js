var Domf = new function () {
	let self = this;
	var $warning;
	var $formSettings;
	var $isNewSite = 0;
	var $isB2B = 0;
	var $delOldMenu = 0;
	var $addDefaultStocktypes = 0;

	this.Init = function() {
		SiteUtils.setTargetDropdownList($('#ddlDOMFSiteList'));
		SiteUtils.Init();

		$warning = $('.warning');

		$formSettings = $('#formDOMFSettings');
		
		/* Event binding */
		$('#ddlDOMFSiteList').on('change', function() {
			SiteUtils.setSiteGuid($(this).val());

			self.getDOMFStatus($('#formDOMFView').eq(0).serialize(), SiteUtils.handleResponse);

			//Update query string with new siteguid parameter
			Location.UpdateURLHash([{"key":"siteguid", "value": $(this).val()}]);
		});

		$('#chkNewSite').on('change', function() {
			$isNewSite = this.checked ? this.value : 0;

			$('#blockWebbizScript').toggleClass('hidden', !this.checked);
		});

		$('#chkB2B').on('change', function() {
			$isB2B = this.checked ? this.value : 0;
		});

		$('#chkDelOldMenu').on('change', function() {
			$delOldMenu = this.checked ? this.value : 0;
		});

		$('#chkStockType').on('change', function() {
			$addDefaultStocktypes = this.checked ? this.value : 0;
		});

		$('input[name=rdScript]:radio').on('change', function() {
			var $val = JSON.parse(this.value);
			if($val) {
				$('#blockOldMenu.hidden').removeClass('hidden');
				$('#blockStockType:not(.hidden)').addClass('hidden');
				$('#chkStockType').prop('checked', false);
			}
			else {
				$('#blockOldMenu:not(.hidden)').addClass('hidden');
				$('#blockStockType.hidden').removeClass('hidden');
				$('#chkDelOldMenu').prop('checked', false);
			}
		});

		$('#btnMakeDOMF').on('click', function() {
			self.proceedMakeDOMF();
		})

		/* e:Event binding */
	}

	this.getDOMFStatus = function(param, callback) {
		SiteUtils.showLoader();
		SiteUtils.hideMessage();

		$.post($proxy, param, function(data, status, xhr) {
			var obj = $.parseJSON(data);

			xPageSystem = "";
			for (var i = 0; i < obj.PageSystem.length; i++) {
				var oPS = obj.PageSystem[i];
				xPageSystem += "<div class='x-guid'>" + oPS.Guid + "</div><div class='x-name'>" + oPS.Name + "</div>";
			}
			$('.gPageSystems').html(xPageSystem);

			xBlockGroup = "";
			for (var i = 0; i < obj.BlockGroup.length; i++) {
				var oBG = obj.BlockGroup[i];
				xBlockGroup += "<div class='x-guid'>" + oBG.Guid + "</div><div class='x-name'>" + oBG.Name + "</div>";
			}
			$('.gBlockGroups').html(xBlockGroup);

			xPageGroup = "";
			for (var i = 0; i < obj.PageGroup.length; i++) {
				var oPG = obj.PageGroup[i];
				xPageGroup += "<div class='x-guid'>" + oPG.Guid + "</div><div class='x-name'>" + oPG.Name + "</div>";
			}
			$('.gPageGroups').html(xPageGroup);

			xBoxTemplate = "";
			for (var i = 0; i < obj.BoxTemplate.length; i++) {
				var oBG = obj.BoxTemplate[i];
				xBoxTemplate += "<div class='x-guid'>" + oBG.Guid + "</div><div class='x-name'>" + oBG.Name + "</div>";
			}
			$('.gBoxTemplate').html(xBoxTemplate);

			xWebsiteSkin = "";
			for (var i = 0; i < obj.WebsiteSkin.length; i++) {
				var oPG = obj.WebsiteSkin[i];
				xWebsiteSkin += "<div class='x-guid'>" + oPG.Guid + "</div><div class='x-name'>" + oPG.Name + "</div>";
			}
			$('.gWebsiteSkin').html(xWebsiteSkin);

			/* should change to Handleresponse ? */
			//self.getAvailableLangAndCurrency({f:'loadLC', siteguid:$('#ddlSiteList').val()});

			SiteUtils.hideLoader();
			SiteUtils.hideMessage();

			if(typeof callback === "function") {
				callback(xhr, data);
			}

		}).fail(function(data) {
			SiteUtils.showResponseMessage(SiteUtils.getErrorMessage(data));
		});
	};

	this.proceedMakeDOMF = function() {
		var $params;

		$params = 	{
						f:'makeDOMF', 
						isNewSite:$isNewSite, 
						isB2B:$isB2B, 
						delOldMenu:$delOldMenu, 
						addDefaultStocktypes:$addDefaultStocktypes, 
						siteguid:SiteUtils.getCurrentSiteGuid()
					};

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
	};

	this.showDOMFStatus = function(responseJSON) {
		let $formInput = $('#formDOMFSettings input');
		if(JSON.parse(responseJSON.IsDOMFSite.toLowerCase())) {
			$formInput.prop('disabled', true);
			$warning.removeClass('green');
		}
		else
		{
			$formInput.prop('disabled', false);
			$warning.addClass('green');
		}
		self.showWarningMessage(responseJSON.MessageText);
	}

	this.showWarningMessage = function(message) {
		$warning.show().html(message);
	}
};

$(document).ready(function() {
	
	// Form initialize
	Domf.Init();
});