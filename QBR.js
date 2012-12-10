function PrintQBR(){
	document.getElementById("btnPrint").style.setAttribute("visibility","hidden");
	parent["main"].print();
	document.getElementById("btnPrint").style.setAttribute("visibility","visible");
}

function openProducts(qbr) {
	alert(qbr);
	//window.open("products.aspx?Nu=" + document.QBR.cache.value + "&Mode=" +
     //   Session("type"),'new','menubar=no,scrollbars=yes,height=450,resizable=yes,width=900,left=100, top=300');
    //alert("in");  
    //window.open('products.aspx');
}

function FormatCurrency(value) {

	while(value.search(',') != -1) {
		if (value.search(',') == value.length - 3) {
			value = value.replace(',','.');
		} else {
			value = value.replace(',','');
		}
	}
		
	if (isNaN(value)) {
		value = 0;
	}
		
	return value;
}
		
function CalculROI() {
	var Cost          = document.QBR.Cost.value == "" ? 0 : FormatCurrency(document.QBR.Cost.value);
	var OnceSavings   = document.QBR.OnceSavings.value == "" ? 0 : FormatCurrency(document.QBR.OnceSavings.value);
	var AnnualSavings = document.QBR.AnnualSavings.value == "" ? 0 : FormatCurrency(document.QBR.AnnualSavings.value);
	
	if (Cost != 0) {
		document.QBR.ROI.value = ((parseFloat(OnceSavings) + parseFloat(AnnualSavings)) / Cost) * 100;
		document.QBR.ROI.value = Math.round(document.QBR.ROI.value * 100) / 100;
	} else {
		document.QBR.ROI.value = 100;
	} 
}
