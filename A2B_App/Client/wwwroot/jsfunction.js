function SampleSelectionReset() {
	document.getElementById("client").selectedIndex = "";
	document.getElementById("sampleRequired").selectedIndex = "";
	document.getElementById("risk").selectedIndex = "";
	document.getElementById("frequency").selectedIndex = "";
	document.getElementById("materialityreq").selectedIndex = "";
	document.getElementById("materialityround1").selectedIndex = "";
	document.getElementById("materialityround2").selectedIndex = "";
	document.getElementById("materialityround3").selectedIndex = "";

	document.getElementById("headerDisplay1").selectedIndex = "";
	document.getElementById("headerDisplay2").selectedIndex = "";
	document.getElementById("headerDisplay3").selectedIndex = "";
	document.getElementById("headerDisplay4").selectedIndex = "";
	document.getElementById("headerDisplay5").selectedIndex = "";

	document.getElementById("headerDisplay21").selectedIndex = "";
	document.getElementById("headerDisplay22").selectedIndex = "";
	document.getElementById("headerDisplay23").selectedIndex = "";
	document.getElementById("headerDisplay24").selectedIndex = "";
	document.getElementById("headerDisplay25").selectedIndex = "";

	document.getElementById("headerDisplay31").selectedIndex = "";
	document.getElementById("headerDisplay32").selectedIndex = "";
	document.getElementById("headerDisplay33").selectedIndex = "";
	document.getElementById("headerDisplay34").selectedIndex = "";
	document.getElementById("headerDisplay35").selectedIndex = "";
}

function SampleSelectionResetFile() {

	document.getElementById("materialityreq").selectedIndex = "";
	document.getElementById("materialityround1").selectedIndex = "";
	document.getElementById("materialityround2").selectedIndex = "";
	document.getElementById("materialityround3").selectedIndex = "";

	document.getElementById("headerDisplay1").selectedIndex = "";
	document.getElementById("headerDisplay2").selectedIndex = "";
	document.getElementById("headerDisplay3").selectedIndex = "";
	document.getElementById("headerDisplay4").selectedIndex = "";
	document.getElementById("headerDisplay5").selectedIndex = "";

	document.getElementById("headerDisplay21").selectedIndex = "";
	document.getElementById("headerDisplay22").selectedIndex = "";
	document.getElementById("headerDisplay23").selectedIndex = "";
	document.getElementById("headerDisplay24").selectedIndex = "";
	document.getElementById("headerDisplay25").selectedIndex = "";

	document.getElementById("headerDisplay31").selectedIndex = "";
	document.getElementById("headerDisplay32").selectedIndex = "";
	document.getElementById("headerDisplay33").selectedIndex = "";
	document.getElementById("headerDisplay34").selectedIndex = "";
	document.getElementById("headerDisplay35").selectedIndex = "";
}

function SetElement(elementId, elementValue) {
	var checkElement = document.getElementById(elementId);
	//this.console.log(`checkElement : '${checkElement}'`);
	if (elementValue !== null && checkElement !== null) {
		if (optionExists(elementValue, document.getElementById(elementId))) {
			document.getElementById(elementId).value = elementValue;
			//this.console.log(`elementId : '${elementId}' - elementValue : '${elementValue}'`);
		}
    }
}


function optionExists(needle, haystack) {
	var optionExists = false,
		optionsLength = haystack.length;

	while (optionsLength--) {
		if (haystack.options[optionsLength].value === needle) {
			optionExists = true;
			break;
		}
	}
	return optionExists;
}


function SetClient(client) {
	document.getElementById("client").value = client;
}

function SetRisk(risk) {
	document.getElementById("risk").value = risk;
}

function SetFrequency(frequency) {
	document.getElementById("frequency").value = frequency;
}

function SetModalClient(client) {
	document.getElementById("client").value = client;
	//document.getElementById("client").selectedIndex = client;
}

window.SetCategoryElementById = function (input) {
	this.console.log(input.id);
	this.console.log(input.client);
	if (input.client) {
		document.getElementById(input.id).value = input.client;
	}
}
function SetCategoryElementById2(id, elemValue) {
	this.console.log(id);
	this.console.log(elemValue);
	if (elemValue) {
		this.console.log(`document.getElementById('${id}').value = '${elemValue}'`);
		document.getElementById(id).value = elemValue;
	}
}

function DownloadFile(path) {
	location.href = path;
}

function DownloadFile2(path) {
	this.console.log(`path = '${path}'`);
	window.open(path);
}

function SaveAsFile(filename) {
	var link = document.createElement('a');
	link.download = filename;
	//link.href = "data:application/octet-stream;base64," + bytesBase64;
	document.body.appendChild(link); // Needed for Firefox
	link.click();
	document.body.removeChild(link);
}

function blazorGetTimezoneOffset() {
	return new Date().getTimezoneOffset();
}