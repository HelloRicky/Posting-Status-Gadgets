
var IMAGE_SOURCE = "url(images/bk1.jpg)";	//background image
var TEXT_PATH = "K:/Eutility/10.0 SmartView/Ricky's/Projects/Windows Gadgets/record/2016_record.txt";	// record path
var googleFormID = "https://docs.google.com/forms/d/1DjSBBJjelB9AxopaHZsuf6HEB29bgw1L0wjoePQFNww/formResponse"
var updateTime = 1000	//in mm sec

// folder location
var parentPath = "L:/Scan Inbox/Energy_Inbox/COMPLETED MONTHLY BILLS/2015";

var Alinta_Foldername = "L:/Scan Inbox/Energy_Inbox/ALINTA";
var Agl_Foldername = "L:/Scan Inbox/Energy_Inbox/AGL ELECTRONIC";
var BGW_Foldername = "L:/Scan Inbox/Energy_Inbox/BGW";
var Mom_Foldername = "L:/Scan Inbox/Energy_Inbox/MOMENTUM ELECTRONIC";
var NextB_Foldername = "L:/Scan Inbox/Energy_Inbox/NextBusiness";
var ERM_Foldername = "L:/Scan Inbox/Energy_Inbox/ERM BILLS";
var ERM_Consol_Foldername = "L:/Scan Inbox/Energy_Inbox/ERM Consolidated Bills";
var EA_Foldername = "L:/Scan Inbox/Energy_Inbox/EA ELECTRONIC";
var EA_Small_Foldername = "L:/Scan Inbox/Energy_Inbox/EA_SMALL MARKET";
var Origin_Foldername = "L:/Scan Inbox/Energy_Inbox/ORIGIN ELECTRONIC";
var Origin_Small_Foldername = "L:/Scan Inbox/Energy_Inbox/ORIGIN SMALL MARKET";
var tcc_Foldername = "L:/Scan Inbox/Energy_Inbox/TCC";

//Retailers array
var AGL_bucket = [Agl_Foldername, "agl"];
var Alinta_bucket = [Alinta_Foldername, "alinta"];
var BGW_bucket = [BGW_Foldername, "bgw"];
var Origin_bucket = [Origin_Foldername, "origin"];
var Origin_Small_bucket = [Origin_Small_Foldername, "origin_small"];
var Momentum_bucket = [Mom_Foldername, "momentum"];
var NextB_bucket = [NextB_Foldername, "nb"];
var ERM_bucket = [ERM_Foldername, "erm"];
var ERM_Consol_bucket = [ERM_Consol_Foldername, "erm_consol"];
var EA_bucket = [EA_Foldername, "ea"];
var EA_Small_bucket = [EA_Small_Foldername, "ea_small"];
var Tcc_bucket = [tcc_Foldername, "tcc"];

//initial
var finalData;
var current_totBill;
var spliter = "$"

//update rate
window.setInterval(refreshGadget, updateTime);

function refreshGadget() {
        location.href = location.href;
}


// Initialize the gadget.
function init()
{
    var oBackground = document.getElementById("imgBackground");
    oBackground.src = IMAGE_SOURCE;
    finalData = "";


	//detect current month
    var d = new Date();
    var n = d.getMonth();
	var localMin = d.getMinutes();
    var m = "";
    var year = d.getUTCFullYear();
	
    if(n==1){
        m = "Jan";
    }else if (n==2){
        m = "Feb";
    }else if (n == 3) {
        m = "Mar";
    }else if (n==4){
        m = "Apr";
    }else if (n==5){
        m = "May";
    }else if (n==6){
        m = "Jun";
    }else if (n==7){
        m = "Jul";
    }else if (n==8){
        m = "Aug";
    }else if (n==9){
        m = "Sep";
    }else if (n==10){
        m = "Oct";
    }else if (n==11){
        m = "Nov";
    }else {
        m = "Dec";
    }

    document.getElementById("header").innerText = "Completed Invoices \n" + m +"-" +year;
	document.getElementById("total_count").innerText = 0	// reset bill no.

	// allocate retailer folders, 
	//ORDER MUST SAME AS SHOWN ON PANEL

	pair_doc(AGL_bucket);
	pair_doc(Origin_bucket);
	pair_doc(Origin_Small_bucket, 2);
	pair_doc(ERM_bucket);	
	pair_doc(ERM_Consol_bucket, 2);	
	pair_doc(Momentum_bucket);
	pair_doc(BGW_bucket);	
	pair_doc(Alinta_bucket);	
	pair_doc(EA_bucket);
	pair_doc(EA_Small_bucket, 2);
	pair_doc(Tcc_bucket);	
	pair_doc(NextB_bucket);	

	current_totBill = document.getElementById("total_count").innerText;

	finalData = current_totBill + spliter + d + finalData;
	recorder(finalData);
	//record data if there is difference
	
	/*if (current_totBill != previous_totBill) {
		document.getElementById("test123").innerText = current_totBill + "- " + previous_totBill
		//recorder(TEXT_PATH, 123);

	}*/
	
   	
}

function pair_doc(file_bucket, offset){
		if (typeof(offset) === 'undefined') offset = 1;
		//get key value pair
		var folderName = file_bucket[0];
		var HTML_id = file_bucket[1];

		var myObject = new ActiveXObject("Scripting.FileSystemObject");
		var fileName = myObject.GetFolder(folderName);
		var fileCount = fileName.files.Count - offset;
		document.getElementById(HTML_id).innerText = fileCount;

		//calculate total bills
		var total_bills = parseInt(document.getElementById("Total_Count").innerText) + fileCount
		document.getElementById("Total_Count").innerText = total_bills
		finalData += (spliter + fileCount);

}

function recorder(data){
	//read file

	var fso, f;
	var words, word;
	var fileWriter, fTest;

	words = "";
	fso = new ActiveXObject("Scripting.FileSystemObject");
	f = fso.OpenTextFile(TEXT_PATH, 1, true);
	
	
	
	//form_data_link(data)
	//httpPost(form_data_link(data))
	
	while (!f.AtEndOfStream){
		
		words = f.ReadLine(); 
		
	}

	if (words.length > 0 ){

		word = words.split(spliter);
		if (word[0] != current_totBill){
			//record in files
			

			fileWriter = new ActiveXObject("Scripting.FileSystemObject");
			fTest = fileWriter.OpenTextFile(TEXT_PATH, 8);
			fTest.WriteLine(data);
			fText.Close;
			fileWriter = nothing;
			fText = nothing;

			//send a copy of data to google doc
			httpPost(form_data_link(data));



		}
		
		
	}



}


function httpPost(theUrl){

	// Post link to google form
    var xmlHttp = new XMLHttpRequest();
    xmlHttp.open( "POST", theUrl, true ); // false for synchronous request
    xmlHttp.send( null );

}

function form_data_link(data_str){
	/*parse data to form a URL*/
	// data_str: 371$Mon Apr 18 15:17:25 UTC+1000 2016$6$0$12$0$1$0$75$40$159$67$0$11
	var data;
	var theUrl = ""
	data = data_str.split(spliter);
	
	var field1 = "entry.1230010703=" + data[2];
	var field2 = "&entry.1001062461=" + data[3];
	var field3 = "&entry.2002109249=" + data[4];
	var field4 = "&entry.1786360573=" + data[5];
	var field5 = "&entry.1640387208=" + data[6];
	var field6 = "&entry.2057825926=" + data[7];
	var field7 = "&entry.505324943=" + data[8];
	var field8 = "&entry.1403124842=" + data[9];
	var field9 = "&entry.1214242439=" + data[10];
	var field10 = "&entry.388647896=" + data[11];
	var field11 = "&entry.1201664961=" + data[12];
	var field12 = "&entry.1857428819=" + data[13];
	var submit = "&submit=submit";
	
	
	theUrl = googleFormID + "?" + field1 + field2 + field3 + field4 + field5 + 
		field6 + field7 + field8 + field9 + field10 + field11 + field12 + submit;
	
	//writeStaff(theUrl)
	return theUrl
}

function writeStaff(data){
	var path_file = "C:/Users/rfzheng/Desktop/123.txt"
	fileWriter = new ActiveXObject("Scripting.FileSystemObject");

			fTest = fileWriter.OpenTextFile(path_file, 8);
			fTest.WriteLine(data);
			fText.Close;
			fileWriter = nothing;
			fText = nothing;

}