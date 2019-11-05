<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
		<title>MCB Accumolo - Go Global maker</title>
		<link rel="stylesheet" href="css/theme.css" type="text/css" />
		<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.12.4/jquery.min.js"></script>
		<script type="text/javascript" src="javascripts/config.js"></script>
		<script type="text/javascript" src="javascripts/utils.js"></script>
		<script type="text/javascript" src="javascripts/goglobal.js"></script>
	</head>
	<body>
		<div class="title">
			<span>Accumolo Go Global - Additional languages and currencies (testing)</span>
		</div>
		<div class="formWrapper">
			<form name="formGoGlobalView" id="formGoGlobalView">
				<input type="hidden" id="formAction" name="f" value="loadData"/>
				<div class="loader"><span>Loading...</span></div>
				<div class="message" id="messageContainer"></div>
				<div class="formTitle"></div>
				<div class="grid">
					<div class="gridRow gridHeader">
						<div class="gridCell">Sites:</div>
						<div class="gridCell wide">
							<select name="siteguid" id="ddlSiteList">
								<option value="0">[Choose your site]</option>
							</select>
						</div>
					</div>
					<div class="gridRow gridHeader">
						<div class="gridCell wide">Added languages:</div>
						<div class="gridCell wide">Added currencies:</div>
					</div>
					<div class="gridRow">
						<div class="gridCell wide gLanguages"></div>
						<div class="gridCell wide gCurrencies"></div>
					</div>
				</div>
			</form>
		</div>
		<div class="formWrapper">
			<form name="formAddData" id="formAddData">
				<input type="hidden" id="formAction" name="f" value="addLC"/>
				<div class="formTitle"></div>
				<div class="grid">
					<div class="gridRow gridHeader">
						<div class="gridCell">Add extra languages and currencies:</div>
					</div>
					<div class="gridRow gridHeader">
						<div class="gridCell">Languages:</div>
						<div class="gridCell wide">
							<select name="sLanguage" id="sLanguage"></select>
							<button type="button" class="x-button" name="btnAddLanguage" id="btnAddLanguage">Add language</button>
						</div>
					</div>
					<div class="gridRow gridHeader">
						<div class="gridCell">Currencies:</div>
						<div class="gridCell wide">
							<select name="sCurrency" id="sCurrency"></select>
							<button type="button" class="x-button" name="btnAddCurrency" id="btnAddCurrency">Add currency</button>
						</div>
					</div>
					<div class="gridRow">
						<div class="gridCell">
							<!--input type="button" class="x-button" id="btnAddLC" value="Add selections"/-->
						</div>
					</div>
				</div>
			</form>
		</div>
		<div class="formWrapper">
			<form name="formCountryFilter" id="formCountryFilter">
				<input type="hidden" id="formAction" name="f" value="fc"/>
				<div class="formTitle"></div>
				<div class="grid">
					<div class="gridRow gridHeader">
						<div class="gridCell">List of country:</div>
					</div>
					<div class="gridRow">
						<div class="gridCell wide">
							<input type="text" placeholder="Enter Country name or GUID or ISO code to filter" name="txtFilter" id="txtFilter">
						</div>
					</div>
				</div>
				<div class="grid" id="tblCountryListHeader">
					<div class="gridRow gridHeader">
						<div class="gridCell cName">Country</div>
						<div class="gridCell">GUID</div>
						<div class="gridCell">IsoCode</div>
						<div class="gridCell">Is EU Member</div>
					</div>
				</div>
				<div class="grid" id="tblCountryList">
				</div>
			</form>
		</div>
	</body>
</html>