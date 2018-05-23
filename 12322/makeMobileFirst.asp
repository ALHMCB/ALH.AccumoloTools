<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
		<title>MCB Accumolo - Mobile First</title>
		<link rel="stylesheet" href="css/theme.css" type="text/css" />
		<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.7/jquery.min.js"></script>
		<script type="text/javascript" src="https://ajax.googleapis.com/ajax/libs/jqueryui/1.8.12/jquery-ui.min.js"></script>
		<script type="text/javascript" src="javascripts/functions.js"></script>
	</head>
	<body>
		<div class="title">
			<span>Accumolo shop - Mobile First</span>
		</div>
		<div class="formWrapper">
			<form name="formGoGlobalView" id="formGoGlobalView">
				<input type="hidden" id="formAction" name="f" value="loadData"/>
				<div class="loader">Loading...</div>
				<div class="message" id="messageContainer"></div>
				<div class="formTitle"></div>
				<div class="grid">
					<div class="gridRow">
						<div class="gridCell">Sites:</div>
						<div class="gridCell wide">
							<select name="siteguid" id="ddlMobileSiteList">
								<option value="0">[Choose your site]</option>
							</select>
						</div>
					</div>
					<div class="gridRow">
						<div class="gridCell wide">Layouts with normal top</div>
					</div>
					<div class="gridRow">
						<div class="gridCell wide gLayouts"></div>
					</div>
				</div>
			</form>
		</div>
		<!-- <div class="formWrapper">
			<form name="formAddData" id="formAddData">
				<input type="hidden" id="formAction" name="f" value="addLC"/>
				<div class="formTitle">Add extra languages and currencies:</div>
				<div class="grid">
					<div class="gridRow">
						<div class="gridCell">Languages:</div>
						<div class="gridCell wide">
							<select name="sLanguage" id="sLanguage"></select>
							<button type="button" class="x-button" name="btnAddLanguage" id="btnAddLanguage">Add language</button>
						</div>
					</div>
					<div class="gridRow">
						<div class="gridCell">Currencies:</div>
						<div class="gridCell wide">
							<select name="sCurrency" id="sCurrency"></select>
							<button type="button" class="x-button" name="btnAddCurrency" id="btnAddCurrency">Add currency</button>
						</div>
					</div>
				</div>
			</form>
		</div> -->
	</body>
</html>