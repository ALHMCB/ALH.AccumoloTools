let PageBase = class {
	constructor(Dropdown, ViewForm, SettingForm, ViewLoader, ViewMessage) {
		this.proxy = "ajax/proxy.asp";
		this.siteGuid = 0;
		this.loader = ViewLoader;
		this.message = ViewMessage;
		this.mainDropdown = Dropdown;
		this.viewForm = ViewForm;
		this.settingForm = SettingForm;

		//const _self = this;

		this.mainDropdown.on('change', () => {
			this.siteGuid = this.mainDropdown.val();

			if(this.siteGuid > 0) {
				this.getSiteStatus(this.viewForm.eq(0).serialize(), PageBaseHelper.handleResponse);
				PageURL.updateURLHash([{"key":"siteguid", "value": this.siteGuid}]);
			}
		})
	}

	getSiteStatus = function(param, handler) {

	}

	getSiteList = function(callback) {
		this.loader.show();
		this.message.hide();

		$.post(this.proxy, {f:'loadSites'}, (data, status, xhr) => {
			let result = $.parseJSON(data);
			let xSites = "";
			result.forEach(function(aSite) {
				xSites += "<option value=" + aSite.Guid + ">" + aSite.Name + "</option>";
			});

			this.mainDropdown.append(xSites);

			this.loader.hide();
			this.message.hide();

			if(typeof callback === "function") {
				callback(xhr, data, this);
			}
		}).fail(function(data) {
			this.showResponseMessage(data);
		});
	}

	// getSiteList = function(callback) {
	// 	this.loader.show();
	// 	this.message.hide();

	// 	const _self = this;

	// 	var result = null;
		
	// 	$.post(this.proxy, {f:'loadSites'}, function(data, status, xhr) {
	// 		result = $.parseJSON(data);

	// 		let xSites = "";
	// 		result.forEach(function(aSite) {
	// 			xSites += "<option value=" + aSite.Guid + ">" + aSite.Name + "</option>";
	// 		});

	// 		_self.mainDropdown.append(xSites);

	// 		_self.loader.hide();
	// 		_self.message.hide();

	// 		if(typeof callback === "function") {
	// 			callback(xhr, data);
	// 		}

	// 	}).fail(function(data) {
	// 		_self.showResponseMessage(data);
	// 	});

	// 	return result;
	// }

	
	static showResponseMessage = function(response) {
		this.loader.hide();
		this.message.show().html(response.MessageText);
	}

	static getErrorMessage = function(data) {
		return {"MessageText":"Error code: " + data.status + ".<br/>Status: " + data.statusText + ".<br/>Response: " + data.responseText};
	}
};

let PageBaseHelper = class {
	static getJSON = function(obj, headerName) {
		if (typeof headerName === 'undefined') {
			headerName = 'X-JSON';
		}
		
		var strJson = obj.getResponseHeader(headerName);
		
		if (typeof obj.getResponseHeader(headerName) === 'string') {
			strJson = strJson.replace(/\\\"/g, '"');
		}
		
		return $.parseJSON(strJson);
	}

	static processSiteSelection = function(entity) {
		if (PageURL.getHash("siteguid") !== "" && $.isNumeric(PageURL.getHash("siteguid"))) {
			var selectedSiteGuid = PageURL.getHash("siteguid");
			
			entity.mainDropdown.val(selectedSiteGuid).change();
		}
	}

	static handleResponse = function(xhr, data, entity = null) {
		const headerJSON = PageBaseHelper.getJSON(xhr);

		if (typeof(headerJSON) === 'object' && headerJSON) {
			switch(headerJSON.EntityType) {
				case "Site":
					PageBaseHelper.processSiteSelection(entity);
					break;
				case "UpdateEntityTable":
					GoGlobal.getLangAndCurrency({f:'loadData', siteguid:_self.siteGuid})
					break;
				case "Domf":
					Domf.showDOMFStatus(headerJSON);
					break;
				case "UpdateDOMFStatus":
					Domf.getDOMFStatus({f:'loadDOMFStatus', siteguid:_self.siteGuid}, _self.handleResponse);
					break;
				case "ShowMessage":
					_self.showResponseMessage(headerJSON.MessageText);
					break;
				default:
					break;
			}
		}
		return $.parseJSON(data);
	}
}