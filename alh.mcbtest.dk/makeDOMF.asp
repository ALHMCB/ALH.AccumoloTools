<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
		<title>MCB DOMF webshop premilinary setup</title>
		<link rel="stylesheet" href="css/theme.css" type="text/css" />
		<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.7/jquery.min.js"></script>
		<script type="text/javascript" src="javascripts/config.js"></script>
		<script type="text/javascript" src="javascripts/utils.js"></script>
		<script type="text/javascript" src="javascripts/domf.js"></script>
		<!-- <script type="text/javascript" src="javascripts/functions.js"></script> -->
	</head>
	<body>
		<div class="title">
			<span>MCB DOMF webshop basic setup</span>
		</div>
		<div class="formWrapper" id="formDOMFViewWrapper">
			<form name="formDOMFView" id="formDOMFView">
				<input type="hidden" id="formAction" name="f" value="loadDOMFStatus"/>
				<div class="loader"><span>Loading...<span></div>
				<div class="message" id="messageContainer"></div>
				<div class="formTitle"></div>
				<div class="grid">
					<div class="gridRow gridHeader">
						<div class="gridCell">Sites:</div>
						<div class="gridCell wide">
							<select name="siteguid" id="ddlDOMFSiteList">
								<option value="0">[Choose your site]</option>
							</select>
						</div>
						<div class="gridCell text wide">
							<label class="warning"></label>
						</div>
					</div>
					<div class="gridRow gridHeader">
						<div class="gridCell wide">Page systems:</div>
						<div class="gridCell wide">Block groups:</div>
						<div class="gridCell wide">Page groups:</div>
						<div class="gridCell wide">Box template group:</div>
						<div class="gridCell wide">Website Skin group:</div>
					</div>
					<div class="gridRow">
						<div class="gridCell wide gPageSystems"></div>
						<div class="gridCell wide gBlockGroups"></div>
						<div class="gridCell wide gPageGroups"></div>
						<div class="gridCell wide gBoxTemplate"></div>
						<div class="gridCell wide gWebsiteSkin"></div>
					</div>
				</div>
			</form>
		</div>
		<div class="formWrapper" id="formDOMFSettingsWrapper">
			<form name="formDOMFSettings" id="formDOMFSettings">
				<input type="hidden" id="formAction" name="f" value="addDOMF" />
				<div class="formTitle"></div>
				<div class="grid">
					<div class="gridRow gridHeader">
						<div class="gridCell">Is new site:</div>
						<div class="gridCell">
							<input type="checkbox" name="chkNewSite" id="chkNewSite" value="1" />
						</div>
					</div>
					<div class="gridRow gridHeader">
						<div class="gridCell">Is B2B:</div>
						<div class="gridCell">
							<input type="checkbox" name="chkB2B" id="chkB2B" value="1" />
						</div>
					</div>
					<div id="blockWebbizScript" class="gridRow gridHeader hidden">
						<div class="gridCell">Did you run MakeWebbiz script?</div>
						<div class="gridCell text narrow flex">
							<input type="radio" name="rdScript" id="rdScript1" value="true"/><label for="rdScript1">Yes</label>
						</div>
						<div class="gridCell text narrow flex">
							<input type="radio" name="rdScript" id="rdScript2" value="false"/><label for="rdScript2">No</label>
						</div>
					</div>
					<div id="blockOldMenu" class="gridRow gridHeader hidden">
						<div class="gridCell">Remove old style menu:</div>
						<div class="gridCell">
							<input type="checkbox" name="chkDelOldMenu" id="chkDelOldMenu" value="1" />
						</div>
					</div>
					<div id="blockStockType" class="gridRow gridHeader hidden">
						<div class="gridCell">Add default stock types:</div>
						<div class="gridCell">
							<input type="checkbox" name="chkStockType" id="chkStockType" value="1" />
						</div>
					</div>
					<div class="gridRow">
						<div class="gridCell">
							<input type="button" class="x-button" name="btnMakeDOMF" id="btnMakeDOMF" value="Proceed"/>
						</div>
					</div>
				</div>
			</form>
		</div>
	</body>
</html>