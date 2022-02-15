var today = new Date();
var todayDate = String(today.getDate()).padStart(2, '0')+'-'+ String(today.getMonth() + 1).padStart(2, '0')+'-'+today.getFullYear();
var formatedTodayDate = today.getFullYear()+'/'+ String(today.getMonth() + 1)+'/'+ String(today.getDate()).padStart(2, '0');

function importToPdfView() {
    var checkPreviousData = $('#element-to-print').html();

    if(checkPreviousData != ''){
        $('#element-to-print').html('');
    }
    
    var regex = /^([a-zA-Z0-9\s_\\.\-:])+(.xlsx|.xls)$/;  
    /*Checks whether the file is a valid excel file*/  
    if (regex.test($("#excelfile").val().toLowerCase())) {  
        var xlsxflag = false; /*Flag for checking whether excel is .xls format or .xlsx format*/  
        if ($("#excelfile").val().toLowerCase().indexOf(".xlsx") > 0) {  
            xlsxflag = true;  
        }  
        /*Checks whether the browser supports HTML5*/  
        if (typeof (FileReader) != "undefined") {  
            var reader = new FileReader();  
            reader.onload = function (e) {  
                var data = e.target.result;  
                /*Converts the excel data in to object*/  
                if (xlsxflag) {  
                    var workbook = XLSX.read(data, { type: 'binary' });  
                }  
                else {  
                    var workbook = XLS.read(data, { type: 'binary' });  
                }  
                /*Gets all the sheetnames of excel in to a variable*/  
                var sheet_name_list = workbook.SheetNames;  
  
                var cnt = 0; /*This is used for restricting the script to consider only first sheet of excel*/  
                sheet_name_list.forEach(function (y) { /*Iterate through all sheets*/  
                    /*Convert the cell value to Json*/  
                    if (xlsxflag) {  
                        var exceljson = XLSX.utils.sheet_to_json(workbook.Sheets[y]);  
                    }  
                    else {  
                        var exceljson = XLS.utils.sheet_to_row_object_array(workbook.Sheets[y]);  
                    }  
                    if (exceljson.length > 0 && cnt == 0) {  
                        BindRow(exceljson);  
                        cnt++;  
                    }
                }); 
            }  
            if (xlsxflag) {/*If excel file is .xlsx extension than creates a Array Buffer from excel*/  
                reader.readAsArrayBuffer($("#excelfile")[0].files[0]);  
            }  
            else {  
                reader.readAsBinaryString($("#excelfile")[0].files[0]);  
            }  
        }  
        else {  
            alert("Sorry! Your browser does not support HTML5!");  
        }  
    }  
    else {  
        alert("Please upload a valid Excel file!");  
    }  
}  

function BindRow(jsondata) {/*Function used to convert the JSON array*/  
    var columns = BindColumnHeader(jsondata); /*Gets all the column headings of Excel*/  

    for (var i = 0; i < jsondata.length; i++) {
        var challanDates = jsondata[i]['Challan Date'].split(',');
        var challanNos = jsondata[i]['Challan No.'].split(',');
        var challanAmounts = jsondata[i]['Challan Amount'].split(',');

        var row = `${i > 0 ? '<div id="element-to-hide" class="py-4" data-html2canvas-ignore="true"><hr></div>' : ''} 
                <div class="single-row">
                    <div class="page-top">
                        <div class="page-header">
                            <table width="100%" style="margin-bottom:20px!important;">
                                <tr>
                                    <td style="border:0px;">Date: ${todayDate}</td>
                                    <td class="text-end" style="border:0px;">Ref: Acc/${formatedTodayDate}-${i+1}</td>
                                </tr>
                            </table>
                        </div>
                        <h3 class="text-center" style="margin-bottom:20px!important;">To Whom It May Concern</h3>
                        This is to certify that Mr ${jsondata[i]['Name']}, ${jsondata[i]['Designation']} of 
                        ${jsondata[i]['Department']} Department is a permanent employee of Devnet Limited. 
                        He joined in the company on ${jsondata[i]['Joining Date']}. The company has paid total Tk  ${jsondata[i]['Total Amount']} (${numberToWords(jsondata[i]['Total Amount'])}) only against as salary and allowance during the financial year ${jsondata[i]['Financial Year']} and assessment year ${jsondata[i]['Assessment Year']}. Details are given below.
                        <br></br>
                    <div>

                    <table width="100%">
                        <thead>
                            <th width="70%">Paticulars</th>
                            <th width="30%" class="text-end">Amount of Taka</th>
                        </thead>
                        <tbody>
                            <tr>
                                <td>Basic Pay</td>
                                <td class="text-end">${jsondata[i]['Basic']}</td>
                            </tr>
                            <tr>
                                <td>House Rent</td>
                                <td class="text-end">${jsondata[i]['Basic']}</td>
                            </tr>
                            <tr>
                                <td>Conveyance</td>
                                <td class="text-end">${jsondata[i]['Conveyance Allowance']}</td>
                            </tr>
                            <tr>
                                <td>Medical</td>
                                <td class="text-end">${jsondata[i]['Medical Allowance']}</td>
                            </tr>
                            <tr>
                                <td>Bonus</td>
                                <td class="text-end">${jsondata[i]['Bonus']}</td>
                            </tr>
                            <tr>
                                <th>Total</th>
                                <th class="text-end">${jsondata[i]['Total Amount']}</th>
                            </tr>
                        </tbody>
                    </table>

                    <div class="mt-3">
                        An amount of Tk. ${jsondata[i]['Total TDS']}/- (${numberToWords(jsondata[i]['Total TDS'])}) only was deducted from his salary as Tax at source under section 50 of the Income Tax Ordinance, 1984.
                    </div>

                    <div class="mt-3">
                        <strong>TDS Payment Details:</strong><br>
                        <table width="100%">
                            <tbody>
                                <tr>
                                    <th>Challan Date</th>
                                    <th>Challan No.</th>
                                    <th>Challan Amount</th>
                                </tr>
                                ${(function fun() {
                                    var challanData = ''; 
                                    for (var n = 0; n < challanDates.length; n++){
                                        challanData += `<tr>
                                                    <td>${challanDates[n]}</td>
                                                    <td>${challanNos[n]}</td>
                                                    <td>${challanAmounts[n]}</td>
                                                </tr>`;
                                    }

                                    return challanData;
                                })()}

                            </tbody>
                        </table>
                    </div>
                    <h5 style="margin-top:100px!important;">Parvin Akhter</h5>
                    Deputy Manager (Accounts & Finance) <br>
                    Devnet Limited
                </div>`;

        $('#element-to-print').append(row);
    }

    $('.download-button').removeClass('d-none');
    $('.element-to-print-div').removeClass('d-none');

    $('html, body').animate({
        scrollTop: $(".pdf-view-devider").offset().top
    }, 500); 
}  

function BindColumnHeader(jsondata) {/*Function used to get all column names from JSON*/  
    var columnSet = [];

    for (var i = 0; i < jsondata.length; i++) {  
        var rowHash = jsondata[i];  
        for (var key in rowHash) {  
            if (rowHash.hasOwnProperty(key)) {  
                if ($.inArray(key, columnSet) == -1) {/*Adding each unique column names to a variable array*/  
                    columnSet.push(key); 
                }  
            }  
        }  
    }  

    return columnSet;  
} 

//Number to word conversion
function numberToWords(number) {  
    var digit = ['zero', 'one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine'];  
    var elevenSeries = ['ten', 'eleven', 'twelve', 'thirteen', 'fourteen', 'fifteen', 'sixteen', 'seventeen', 'eighteen', 'nineteen'];  
    var countingByTens = ['twenty', 'thirty', 'forty', 'fifty', 'sixty', 'seventy', 'eighty', 'ninety'];  
    var shortScale = ['', 'thousand', 'million', 'billion', 'trillion'];  

    number = number.toString(); number = number.replace(/[\, ]/g, ''); if (number != parseFloat(number)) return 'not a number'; var x = number.indexOf('.'); if (x == -1) x = number.length; if (x > 15) return 'too big'; var n = number.split(''); var str = ''; var sk = 0; for (var i = 0; i < x; i++) { if ((x - i) % 3 == 2) { if (n[i] == '1') { str += elevenSeries[Number(n[i + 1])] + ' '; i++; sk = 1; } else if (n[i] != 0) { str += countingByTens[n[i] - 2] + ' '; sk = 1; } } else if (n[i] != 0) { str += digit[n[i]] + ' '; if ((x - i) % 3 == 0) str += 'hundred '; sk = 1; } if ((x - i) % 3 == 1) { if (sk) str += shortScale[(x - i - 1) / 3] + ' '; sk = 0; } } if (x != number.length) { var y = number.length; str += 'point '; for (var i = x + 1; i < y; i++) str += digit[n[i]] + ' '; } str = str.replace(/\number+/g, ' '); return str.trim();
}

//for download pdf
function downloadPdf() {
    var element = document.getElementById('element-to-print');
    var opt = {
        margin: 0.5,
        filename: 'myfile.pdf',
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 2},
        jsPDF: { unit: 'in', format: 'letter', orientation: 'portrait' },
        pagebreak: { mode: 'avoid-all', after: '.single-row' }
    };

    // New Promise-based usage:
    html2pdf().set(opt).from(element).save();
}

//for scroll to top
$(document).ready(function(){ 
    $(window).scroll(function(){ 
        if ($(this).scrollTop() > 500) { 
            $('#page-scroll').fadeIn(); 
        } else { 
            $('#page-scroll').fadeOut(); 
        } 
    }); 

    $('#page-scroll').click(function(){ 
        $("html, body").animate({ 
            scrollTop: $(".pdf-view-devider").offset().top
        }, 600); 
        return false; 
    }); 
});