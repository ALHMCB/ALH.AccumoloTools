let PageDOMF = class extends PageBase {
	constructor(Dropdown, ViewForm, SettingForm, ViewLoader, ViewMessage, ViewWarning) {
		super(Dropdown, ViewForm, SettingForm, ViewLoader, ViewMessage);
		this.warning = ViewWarning;
	}

	getSiteStatus = function(param, handler) {
		this.loader.show();
		this.message.hide();

		$.post(this.proxy, param, (data, status, xhr) => {
			const jsonData = $.parseJSON(data);

			let xPageSystem = "";
			jsonData.PageSystem.forEach(obj => {
				xPageSystem += "<div class='x-guid'>" + obj.Guid + "</div><div class='x-name'>" + obj.Name + "</div>";
			});
			$('.gPageSystems').html(xPageSystem);

			let xBlockGroup = "";
			jsonData.BlockGroup.forEach(obj => {
				xBlockGroup += "<div class='x-guid'>" + obj.Guid + "</div><div class='x-name'>" + obj.Name + "</div>";
			});
			$('.gBlockGroups').html(xBlockGroup);

			let xPageGroup = "";
			jsonData.PageGroup.forEach(obj => {
				xBlockGroup += "<div class='x-guid'>" + obj.Guid + "</div><div class='x-name'>" + obj.Name + "</div>";
			});
			$('.gPageGroups').html(xPageGroup);

			let xBoxTemplate = "";
			jsonData.BoxTemplate.forEach(obj => {
				xBlockGroup += "<div class='x-guid'>" + obj.Guid + "</div><div class='x-name'>" + obj.Name + "</div>";
			});
			$('.gBoxTemplate').html(xBoxTemplate);

			let xWebsiteSkin = "";
			jsonData.WebsiteSkin.forEach(obj => {
				xBlockGroup += "<div class='x-guid'>" + obj.Guid + "</div><div class='x-name'>" + obj.Name + "</div>";
			});
			$('.gWebsiteSkin').html(xWebsiteSkin);

			/* should change to Handleresponse ? */
			//self.getAvailableLangAndCurrency({f:'loadLC', siteguid:$('#ddlSiteList').val()});

			this.loader.hide();
			this.message.hide();

			if(typeof handler === "function") {
				handler(xhr, data);
			}

		}).fail(function(data) {
			this.showResponseMessage(this.getErrorMessage(data));
		});
	}
}

