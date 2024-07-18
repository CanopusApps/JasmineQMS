$(document).ready(function () {
    //alert('Hello in new js file');
});

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
function ToJavaScriptDate(value) {
    if (value == null) return;
    var pattern = /Date\(([^)]+)\)/;
    var results = pattern.exec(value);
    var dt = new Date(parseFloat(results[1]));
    return (dt.getMonth() + 1) + "/" + dt.getDate() + "/" +  dt.getFullYear();
}
function ToJavaScriptDateWithFormat(value) {
    if (value == null) return;
    var pattern = /Date\(([^)]+)\)/;
    var results = pattern.exec(value);
    var dt = new Date(parseFloat(results[1]));
    return dt.getDate() + "-" + (dt.getMonth() + 1) + "-" +  dt.getFullYear();
}

function ShowProgressBar() {
    $("#overlay").show();
    $("#overlay").addClass('modal-backdrop fade in');
    $("#sidebar").addClass('modal-backdrop fade in');
}

function HideProgressBar() {
    $("#overlay").hide();
    $("#overlay").removeClass('modal-backdrop fade in');
    $("#sidebar").removeClass('modal-backdrop fade in');
}

function ValidateEmail(email) {
    var expr = /^([\w-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([\w-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$/;
    return expr.test(email);
};

function IsFileTypeAllowed(ftypes, fext) {
    var IsAllowed = false;
    var result = ftypes.indexOf(fext);
    if (result != -1) {
        IsAllowed = true;
    }
    else {
        IsAllowed = false;
        $('.btn-warning').trigger('click');
        var msg = 'File Type <b>' + fext + '</b> is not supported. </br> Allowed File Types are <b>' + ftypes + ' </b>';
        $('.modal-body').html(msg);
    }
    return IsAllowed;
}

function IsFileSizeAllowed(fileSize, allowedSize) {
    var IsAllowed = false;
    if (fileSize == 0) return false;
    var sizeinKB = fileSize / 1024;
    var sizeinMB = sizeinKB / 1024;
    if (sizeinMB > allowedSize) {
        IsAllowed = false;
        $('.btn-warning').trigger('click');
        var msg = 'File size should not be more than ' + allowedSize + ' MB';
        $('.modal-body').html(msg);
    }
    else {
        IsAllowed = true;
    }
    return IsAllowed;
}

function ValidateFormFields(fieldset) {
    var ermsg = '';
    $(".mantable").find(fieldset).each(function () {
        var type = $(this).prop("type");
        var fieldName = $(this).prop("title");
        if (type == 'text') {
            if ($(this).val().trim() == '') {
                ermsg += fieldName + '<br />';
                $(this).addClass("manfieldredborder");
            }
            else {
                $(this).removeClass("manfieldredborder");
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
        else if (type == 'select-one') {
            if ($(this).val() == '0') {
                ermsg += fieldName + '<br />';
                $(this).addClass("manfieldredborder");
            }
            else
                $(this).removeClass("manfieldredborder");
        }
        else if (type == 'file') {
            if ($(this).val() == '') {
                ermsg += fieldName + '<br />';
                $(this).addClass("manfieldredborder");
            }
            else
                $(this).removeClass("manfieldredborder");
        }
    });
    //if ($("#date1") != undefined) {
    //    if ($("#date1").val().trim() != "" && !IsDateLessThanToDay($("#date1").val().trim())) {
    //        ermsg += "Date of Birth should be less than today.";
    //        $("#date1").addClass("manfieldredborder");
    //    }
    //    else
    //        $("#date1").removeClass("manfieldredborder");
    //}
    if (ermsg != '') {
        $('.btn-warning').trigger('click');
        $('.modal-body').html(ermsg);
        return false;
    }
    return true;
}

function GenerateDocumentNumber(cururl, ctrl, projCode, deptCode, secCode, catCode, guidctrl, executionctrl) {
    $.ajax({
        url: cururl,
        dataType: "json",
        type: "POST",
        contentType: 'application/json; charset=utf-8',
        beforeSend: ShowProgressBar(),
        complete: HideProgressBar(),
        cache: false,
        async: true,
        data: JSON.stringify({ projCode: projCode, deptCode: deptCode, secCode: secCode, catCode: catCode }),
        success: function (data) {
            if (data.success) {
                var msg = data.message1.toString().split('#');
                $(ctrl).val(msg[0]);
                $(guidctrl).text(msg[1]);
                $(executionctrl).text(msg[2]);
            }
        },
        error: function (xhr) {
            alert('error ' + xhr.message);
        }
    });
}


function BindSectionsForDept(cururl, $ctrl1, deptID, defaultOption) {
    $.ajax({
        url: cururl,
        dataType: "json",
        type: "POST",
        contentType: 'application/json; charset=utf-8',
        beforeSend: ShowProgressBar(),
        complete: HideProgressBar(),
        cache: false,
        async: true,
        data: JSON.stringify({ deptID: deptID }),
        success: function (data) {
            if (data.success) {
                BindDropdown(data.message1, $ctrl1, defaultOption);
            }
        },
        error: function (xhr) {
            alert('error ' + xhr.message);
        }
    });
}
function BindStagesForWorkflow(cururl, $ctrl1, workflowID, defaultOption) {
    $.ajax({
        url: cururl,
        dataType: "json",
        type: "POST",
        contentType: 'application/json; charset=utf-8',
        beforeSend: ShowProgressBar(),
        complete: HideProgressBar(),
        cache: false,
        async: true,
        data: JSON.stringify({ wfID: workflowID }),
        success: function (data) {
            if (data.success) {
                BindDropdown(data.message1, $ctrl1, defaultOption);
            }
        },
        error: function (xhr) {
            alert('error ' + xhr.message);
        }
    });
}
function BindDropdown(data, ctrl, defaultOption) {
    ctrl.find("option").remove();
    ctrl.append('<option title="' + defaultOption + '" value="0">' + defaultOption+'</option>');
    for (var i = 0; i < data.length; i++) {
        var code = ''; var wfid = '', Condition = ''; var projTypeCode = ''; var documentLevel = '';
        if (data[i].Code != undefined && data[i].Code != null) {
            code = data[i].Code;
        }
        if (data[i].Condition != undefined && data[i].Condition != null) {
            Condition = data[i].Condition;
        }
        if (data[i].WorkflowID != undefined && data[i].WorkflowID != null) {
            wfid = data[i].WorkflowID;
        }
        if (data[i].ProjectTypeCode != undefined && data[i].ProjectTypeCode != null) {
            projTypeCode = data[i].ProjectTypeCode;
        }
        if (data[i].DocumentLevel != undefined && data[i].DocumentLevel != null) {
            documentLevel = data[i].DocumentLevel;
        }
        ctrl.append('<option documentLevel="' + documentLevel +'" projType="' + data[i].ProjectType+'" projTypeCode="' + projTypeCode +'" projTypeID="' + data[i].ProjectTypeID +'" workflowID = "' + wfid + '" code="' + code + '" condition="' + Condition + '" title="' + data[i].Title + '" value="' + data[i].ID + '">' + data[i].Title + '</option>');
    }
}
function BindDropDownBoxes(cururl, $ctrl1, $ctrl2, $ctrl3, $ctrl4, defaultOption, IsProjectActive) {
    $.ajax({
        url: cururl,
        dataType: "json",
        type: "POST",
        contentType: 'application/json; charset=utf-8',
        beforeSend: ShowProgressBar(),
        complete: HideProgressBar(),
        cache: false,
        async: true,
        data: JSON.stringify({ IsProjectActive: IsProjectActive }),
        success: function (data) {
            if (data.success) {
                if ($ctrl1 != '')
                    BindDropdown(data.message1, $ctrl1, defaultOption);
                if ($ctrl2 != '')
                    BindDropdown(data.message2, $ctrl2, defaultOption);
                if ($ctrl3 != '')
                    BindDropdown(data.message3, $ctrl3, defaultOption);
                if ($ctrl4 != '')
                    BindDropdown(data.message4, $ctrl4, defaultOption);
            }
        },
        error: function (xhr) {
            alert('error ' + xhr.message);
        }
    });
}
function BindDropDownBoxes_Admin(cururl, $ctrl1, $ctrl2, $ctrl3, $ctrl4, ctrl5, ctrl6, defaultOption) {
    $.ajax({
        url: cururl,
        dataType: "json",
        type: "POST",
        contentType: 'application/json; charset=utf-8',
        beforeSend: ShowProgressBar(),
        complete: HideProgressBar(),
        cache: false,
        async: true,
        //data: JSON.stringify({ countryID: cID }),
        success: function (data) {
            if (data.success) {
                if ($ctrl1 != '')
                    BindDropdown(data.message1, $ctrl1, defaultOption);
                if ($ctrl2 != '')
                    BindDropdown(data.message2, $ctrl2, defaultOption);
                if ($ctrl3 != '')
                    BindDropdown(data.message1, $ctrl3, defaultOption);
                if ($ctrl4 != '')
                    BindDropdown(data.message2, $ctrl4, defaultOption);
                if (ctrl5 != '')
                    BindDropdown(data.message3, ctrl5, defaultOption);
                if (ctrl6 != '')
                    BindDropdown(data.message3, ctrl6, defaultOption);
            }
        },
        error: function (xhr) {
            alert('error ' + xhr.message);
        }
    });
}

function BindDropDownBoxes_AdminandUser(cururl, $ctrl1, $ctrl2, $ctrl3, $ctrl4, ctrl5, ctrl6, defaultOption, isReadOnly) {
    $.ajax({
        url: cururl,
        dataType: "json",
        type: "POST",
        contentType: 'application/json; charset=utf-8',
        beforeSend: ShowProgressBar(),
        complete: HideProgressBar(),
        cache: false,
        async: true,
        data: JSON.stringify({ isReadOnly: isReadOnly }),
        success: function (data) {
            if (data.success) {
                if ($ctrl1 != '')
                    BindDropdown(data.message1, $ctrl1, defaultOption);
                if ($ctrl2 != '')
                    BindDropdown(data.message2, $ctrl2, defaultOption);
                if ($ctrl3 != '')
                    BindDropdown(data.message1, $ctrl3, defaultOption);
                if ($ctrl4 != '')
                    BindDropdown(data.message2, $ctrl4, defaultOption);
                if (ctrl5 != '')
                    BindDropdown(data.message3, ctrl5, defaultOption);
                if (ctrl6 != '')
                    BindDropdown(data.message3, ctrl6, defaultOption);
            }
        },
        error: function (xhr) {
            alert('error ' + xhr.message);
        }
    });
}

function BindProjStagesForPType(cururl, $ctrl1, $ctrl2,projTypeID, workflowID, defaultOption) {
    $.ajax({
        url: cururl,
        dataType: "json",
        type: "POST",
        contentType: 'application/json; charset=utf-8',
        beforeSend: ShowProgressBar(),
        complete: HideProgressBar(),
        cache: false,
        async: true,
        data: JSON.stringify({ projTypeID: projTypeID, wfID: workflowID }),
        success: function (data) {
            if (data.success) {
                if ($ctrl1 != '')
                    BindDropdown(data.message1, $ctrl1, defaultOption);
                if ($ctrl2 != '')
                    BindDropdown(data.message2, $ctrl2, defaultOption);
            }
        },
        error: function (xhr) {
            alert('error ' + xhr.message);
        }
    });
}

function BindProjStagesForPType_2(cururl, $ctrl1, $ctrl2, projTypeID, workflowID, defaultOption, isReadOnly) {
    $.ajax({
        url: cururl,
        dataType: "json",
        type: "POST",
        contentType: 'application/json; charset=utf-8',
        beforeSend: ShowProgressBar(),
        complete: HideProgressBar(),
        cache: false,
        async: true,
        data: JSON.stringify({ projTypeID: projTypeID, wfID: workflowID, isReadOnly: isReadOnly }),
        success: function (data) {
            if (data.success) {
                if ($ctrl1 != '')
                    BindDropdown(data.message1, $ctrl1, defaultOption);
                if ($ctrl2 != '')
                    BindDropdown(data.message2, $ctrl2, defaultOption);
            }
        },
        error: function (xhr) {
            alert('error ' + xhr.message);
        }
    });
}

function BindSubCatsForCategory(cururl, $ctrl1, catID, defaultOption) {
    $.ajax({
        url: cururl,
        dataType: "json",
        type: "POST",
        contentType: 'application/json; charset=utf-8',
        beforeSend: ShowProgressBar(),
        complete: HideProgressBar(),
        cache: false,
        async: false,
        data: JSON.stringify({ catID: catID }),
        success: function (data) {
            if (data.success) {
                BindDropdown(data.message1, $ctrl1, defaultOption);
            }
        },
        error: function (xhr) {
            alert('error ' + xhr.message);
        }
    });
}



