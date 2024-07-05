var dateFormat = "MM/dd/yyyy";

$(document).ready(function () {
    //alert('Hello in new js file');
});


function DeleteUser(id)
{
    var cururl = "/TestUser/DeleteEntity";
    $.ajax({
        url: cururl,
        dataType: "json",
        type: "POST",
        beforeSend: function () {
            $("#overlay").show();
            $("#overlay").addClass('modal-backdrop fade in');
            $("#sidebar").addClass('modal-backdrop fade in');
        },
        complete: function () {
            $("#overlay").hide();
            $("#overlay").removeClass('modal-backdrop fade in');
            $("#sidebar").removeClass('modal-backdrop fade in');
        },
        contentType: 'application/json; charset=utf-8', //define a contentType of your request
        cache: false, //avoid caching results
        data: JSON.stringify({ id: id }), // here you can pass arguments to your request if you need
        success: function (data) {
            if (data.success) {
                $('.btn-success').trigger('click');
                $('.modal-body').html('User Deleted Successfully.');
                retmsg = data.message;
            }
        },
        error: function (xhr) {
            $('.btn-danger').trigger('click');
            $('.modal-body').html('Error while deleting User.');
        }
    });
}
function SaveUserToDB(id,name,em,ph)
{
    //alert('Save User To DB');
    var lst = new Array();
    lst[0] = id.trim();
    lst[1] = name.trim();
    lst[2] = em.trim();
    lst[3] = ph.trim();
    var retmsg = '';
    var cururl = "/TestUser/SaveEntity";
    $.ajax({
        url: cururl,
        dataType: "json",
        type: "POST",
        beforeSend: function () {
            $("#overlay").show();
            $("#overlay").addClass('modal-backdrop fade in');
            $("#sidebar").addClass('modal-backdrop fade in');
        },
        complete: function () {
            $("#overlay").hide();
            $("#overlay").removeClass('modal-backdrop fade in');
            $("#sidebar").removeClass('modal-backdrop fade in');
        },
        async: false,
        contentType: 'application/json; charset=utf-8', //define a contentType of your request
        cache: false, //avoid caching results
        data: JSON.stringify({ arr: lst }), // here you can pass arguments to your request if you need
        success: function (data) {
            if (data.success) {
                //$('.btn-success').trigger('click');
                //$('.modal-body').html('Employee Saved Successfully.');
                //retmsg = data.message;
                //alert('in ajax' + retmsg);
                retmsg = data.message;
                var mesge = '';
                if (retmsg == '0' || retmsg.indexOf("#") >= 0) {
                    $('.btn-warning').trigger('click');
                    $('.modal-body').html('User with same email is already exists.');
                }
                else {
                    $('.btn-success').trigger('click');
                    $('.modal-body').html('User Saved Successfully.');

                }

            }
        },
        error: function (xhr) {
            $('.btn-danger').trigger('click');
            $('.modal-body').html('Error while saving User.');
        }
    });
    return retmsg;
}
function checkForUserSave(nm, em, ph) {
    //alert('hi');
    //alert(nm + '-' + ph);
    var ermsg = '';
    if (nm.trim() == '') {
        ermsg = 'Name<br/>';
    }
    if (em.trim() == '') {
        ermsg += 'Email<br/>';
    }
    if (ph.trim() == '') {
        ermsg += 'Phone<br/>';
    }

    if (em.trim() != '' && !ValidateEmail(em.trim())) {
        ermsg += 'Email is not Valid.<br/>';
    }

    if (ermsg != '') {
        $('.btn-warning').trigger('click');
        $('.modal-body').html(ermsg);
        return false;
    }
    return true;
}





//App User - START
function BindControlMultipleItems(rIDs, cururl, $ctrl) {
    $.ajax({
        url: cururl,
        dataType: "json",
        type: "POST",
        contentType: 'application/json; charset=utf-8',
        beforeSend: ShowProgressBar(),
        complete: HideProgressBar(),
        cache: false,
        async:true,
        //data: JSON.stringify({ countryID: cID }),
        success: function (data) {
            if (data.success) {
                var optiongrp = '';
                //$('<optGroup/>').attr('label', 'Roles').appendTo($cont);
                $ctrl.find("option").remove();
                for (var i = 0; i < data.message.length; i++) {
                    //if (i == 0 || optiongrp != data.message[i].CreatedBy) {
                    //    optiongrp = data.message[i].CreatedBy;
                    //    $('<optGroup/>').attr('label', optiongrp).appendTo($cont).append('<option title="' + data.message[i].Description + '" value="' + data.message[i].ID + '">' + data.message[i].Title + '</option>');
                    //} else {
                    //    $cont.find('optGroup').last().append('<option title="' + data.message[i].Description + '" value="' + data.message[i].ID + '">' + data.message[i].Title + '</option>');                        
                    //}
                    //debugger;
                    //alert(data.message[i]);

                    $ctrl.append('<option title="' + data.message[i].Description + '" value="' + data.message[i].ID + '">' + data.message[i].Description + '</option>');
                }

                $ctrl.multipleSelect('refresh');
                //alert(rIDs);
                if (rIDs != '')
                    $ctrl.multipleSelect("setSelects", rIDs);
            }
        },
        error: function (xhr) {
            alert('error');
        }
    });
}

function SaveEntitytoDB(oTable,cururl, id, arrData, arry2, newRow,entity)
{
    var retmsg = '';
    var svText = "Saved";
    if (id != "")
        svText = "Updated";
    $.ajax({
        url: cururl,
        dataType: "json",
        type: "POST",
        beforeSend: ShowProgressBar(),
        complete: HideProgressBar(),
        async: false,
        contentType: 'application/json; charset=utf-8',
        cache: false,
        data: JSON.stringify({ arr: arrData, arr2: arry2 }),
        success: function (data) {
            if (data.success) {
                retmsg = data.message;
                var mesge = '';
                if (retmsg == '0' || retmsg.indexOf("#") >= 0) {
                    $('.btn-warning').trigger('click');
                    $('.modal-body').html(entity + ' with same email is already exists.');

                }
                else if (retmsg == '-1') {
                    $('.btn-danger').trigger('click');
                    $('.modal-body').html('Error while saving ' + entity + '.');
                }
                else {
                    $('.btn-success').trigger('click');
                    $('.modal-body').html(entity + ' ' + svText + ' Successfully.');

                    //create
                    if (id == "") {
            //            nRow = oTable.fnAddData(['-1', arrData[1], arrData[4], arrData[2], arrData[5],
            //'<a href="#" class="btn btn-info btn-xs edit"><i class="fa fa-edit"></i> Edit</a><a href="#" class="btn btn-danger btn-xs delete"><i class="fa fa-trash-o"></i> Delete</a>'
                        //            ]);
                        var nRow = oTable.fnAddData(newRow);
                        oTable.fnUpdate(retmsg, nRow, 0, false);
                        resetAllValues();
                    }
                    else {
                        //update
                        resetAllValues();
                        $("#divnew").hide();
                    }
                }
            }
        },
        error: function (xhr) {
            $('.btn-danger').trigger('click');
            $('.modal-body').html('Error while saving ' + entity + '.');
        }
    });
    return retmsg;
}

function DeleteEntityItem(id, cururl,entity) {
    $.ajax({
        url: cururl,
        dataType: "json",
        type: "POST",
        beforeSend: ShowProgressBar(),
        complete: HideProgressBar(),
        contentType: 'application/json; charset=utf-8',
        cache: false,
        data: JSON.stringify({ id: id }),
        success: function (data) {
            if (data.success) {
                $('.btn-success').trigger('click');
                $('.modal-body').html(entity + ' Deleted Successfully.');
                retmsg = data.message;
            }
            else
            {
                $('.btn-danger').trigger('click');
                $('.modal-body').html('Error while deleting ' + entity + '.');
            }
        },
        error: function (xhr) {
            $('.btn-danger').trigger('click');
            $('.modal-body').html('Error while deleting ' + entity + '.');
        }
    });
}

//AppUser - END

//Common Functions
function ValidateFormFields() {
    var ermsg = '';
    $(".mantable").find(".manfield").each(function () {
        //alert($(this).val());
        var type = $(this).prop("type");
        var fieldName = $(this).prop("title");

        if (type == 'text') {
            if ($(this).val().trim() == '') {
                ermsg += fieldName + '<br />';
                $(this).addClass("manfieldredborder");
            }
            else {
                $(this).removeClass("manfieldredborder");

                //To validate the email id
                if ($(this).prop("class").indexOf("propemail") != -1) {
                    if (!ValidateEmail($(this).val().trim())) {
                        ermsg += 'Email ID is not valid.<br />';
                        $(this).addClass("manfieldredborder");
                    }
                    else
                        $(this).removeClass("manfieldredborder");
                }
            }
        }
    });

    if ($("#date1").val().trim() != "" && !IsDateLessThanToDay($("#date1").val().trim()))
    {
        ermsg += "Date of Birth should be less than today.";
        $("#date1").addClass("manfieldredborder");
    }
    else
        $("#date1").removeClass("manfieldredborder");


    if (ermsg != '') {
        $('.btn-warning').trigger('click');
        $('.modal-body').html(ermsg);
        return false;
    }
    return true;
}

function getAllFormValues() {
    var inputValues = new Array();
    inputValues = $('.mantable :input').map(function () {
        var type = $(this).prop("type");
           
        if (type == "checkbox") {
            if ($(this).prop("id").trim() != "") {
                if (this.checked) {
                    return "checked";
                }
                else
                    return "notchecked";
            }
        }
        else if ((type == "radio")) {
            if (this.checked)
                return "yes";
            else
                return "no";
        }
        else if (type != "button" && type != "submit" && type != "select-multiple") {
            return $(this).val().trim();
        }
    })
    return inputValues;
}

function resetAllValues() {
    $('.mantable').find('input:text').val('');
    $("#ddlMultiple").multipleSelect("setSelects", '');
    $("#rbn1").prop("checked", true);
    $("#chk1").prop('checked', true);
    $("#divID").html('');
}

function GetIDsFromMultipleSelected(ctrl) {
    //alert($(ctrl + " option:selected").length);
    var arr = new Array(); var i = 0;
    if ($(ctrl + " option:selected").length > 0) {
        $(ctrl + " option:selected").each(function () {
            //debugger;
            var $this = $(this);
            //alert($this.length);
            if ($this.length) {
                var selText = $this.text();
                var selValue = $this.val();
                i = i + 1;
                arr[i] = selValue;
            }
        });
    }
    return arr;
}

function ValidateEmail(email) {
    var expr = /^([\w-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([\w-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$/;
    return expr.test(email);
};

function AllowNumbers(evtobj) {
    //alert(evtobj.keyCode);
    // Allow: backspace, delete, tab, escape, enter and 
    if ($.inArray(evtobj.keyCode, [46, 8, 9, 27, 13, 110, 188]) !== -1 ||
        // Allow: Ctrl+A, Command+A -- 86 is for 'V' key
        ((evtobj.keyCode == 65 || evtobj.keyCode == 86 || evtobj.keyCode == 16) && (evtobj.ctrlKey === true || evtobj.metaKey === true)) ||
        // Allow: home, end, left, right, down, up
        (evtobj.keyCode >= 35 && evtobj.keyCode <= 40)) {
        // let it happen, don't do anything
        return;
    }
    //alert($('#txtPhone').length);
    // Ensure that it is a number and stop the keypress
    if ((evtobj.shiftKey || (evtobj.keyCode < 48 || evtobj.keyCode > 57)) && (evtobj.keyCode < 96 || evtobj.keyCode > 105)) {
        evtobj.preventDefault();
    }
}

function IsDateLessThanToDay(date1) {
    var valid = true;
    var sDateStr = date1;
    var sDateArr = sDateStr.split("-");
    var cuFormat = dateFormat.split("-");
    var sDate;
    if (cuFormat[0] == "MM")
        sDate = new Date(sDateArr[2], sDateArr[0] - 1, sDateArr[1]);
    else
        sDate = new Date(sDateArr[2], sDateArr[1] - 1, sDateArr[0]);
    var todayDate = new Date();
    todayDate.setHours(0, 0, 0, 0);

    if (todayDate <= sDate) {
        valid = false;
    };
    return valid;
}

function validEDate(date1, date2) {
    var valid = true;
    var sDateArr = date1.split("-");
    var cuFormat = dateFormat.split("-");
    var sDate; var eDate;
    if (cuFormat[0] == "MM")
        sDate = new Date(sDateArr[2], sDateArr[0] - 1, sDateArr[1]);
    else
        sDate = new Date(sDateArr[2], sDateArr[1] - 1, sDateArr[0]);

    var sDate2 = date2.split("-");
    if (cuFormat[0] == "MM")
        eDate = new Date(sDate2[2], sDate2[0] - 1, sDate2[1]);
    else
        eDate = new Date(sDate2[2], sDate2[1] - 1, sDate2[0]);

    if (sDate > eDate) {
        valid = false;
    };
    return valid;
}

function DateLessThanOREqToDay(date1) {

    var valid = true;
    var sDateStr = date1;
    var sDateArr = sDateStr.split("-");
    var cuFormat = dateFormat.split("-");
    var sDate;
    if (cuFormat[0] == "MM")
        sDate = new Date(sDateArr[2], sDateArr[0] - 1, sDateArr[1]);
    else
        sDate = new Date(sDateArr[2], sDateArr[1] - 1, sDateArr[0]);
    var todayDate = new Date();
    todayDate.setHours(0, 0, 0, 0);

    if (todayDate < sDate) {
        valid = false;
    };
    return valid;
}

function ValidateFileExtension(fileC) {
    var sFileName = fileC.val();
    if (sFileName.length > 0) {
        var blnValid = false;
        for (var j = 0; j < _validFileExtensions.length; j++) {
            var sCurExtension = _validFileExtensions[j];
            if (sFileName.substr(sFileName.length - sCurExtension.length, sCurExtension.length).toLowerCase() == sCurExtension.toLowerCase()) {
                blnValid = true;
                break;
            }
        }

        if (!blnValid) {
            //alert("Sorry, " + sFileName + " is invalid, allowed extensions are: " + _validFileExtensions.join(", "));
            return false;
        }
    }
    return true;
}

function ToJavaScriptDate(value) {
    var pattern = /Date\(([^)]+)\)/;
    var results = pattern.exec(value);
    var dt = new Date(parseFloat(results[1]));
    return dt.getDate() + "-" + (dt.getMonth() + 1) + "-" + dt.getFullYear();
}

function ShowProgressBar()
{
    //alert('ShowProgressBar');
    $("#overlay").show();
    $("#overlay").addClass('modal-backdrop fade in');
    $("#sidebar").addClass('modal-backdrop fade in');
}

function HideProgressBar() {
    //alert('HideProgressBar');
    $("#overlay").hide();
    $("#overlay").removeClass('modal-backdrop fade in');
    $("#sidebar").removeClass('modal-backdrop fade in');
}
