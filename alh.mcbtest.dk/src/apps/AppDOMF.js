var appDOMF;
$(document).ready(function() {
	appDOMF = new PageDOMF(
		$('#ddlDOMFSiteList'), 
		$('#formDOMFView'), 
		$('#formDOMFSettings'),
		$('.loader'),
		$('.message'),
		$('.warning')
	);

	appDOMF.getSiteList(PageDOMF.handleResponse);
});