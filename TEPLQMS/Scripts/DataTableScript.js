
var InitiateEditableDataTable = function () {
    return {
        init: function () {
            //Datatable Initiating
            var oTable = $('#editabledatatable').dataTable({
                "aLengthMenu": [
                    [5, 10, 20, 100, -1],
                    [5, 10, 20, 100, "All"]
                ],
                "iDisplayLength": 10,
                "sPaginationType": "bootstrap",
                "sDom": "Tflt<'row DTTTFooter'<'col-sm-6'i><'col-sm-6'p>>",
                "oTableTools": {
                    "aButtons": [
                        //"copy",
                        //"print",                           
                        //{
                        //    "sExtends": "collection",
                        //    "sButtonText": "Save <i class=\"fa fa-angle-down\"></i>",
                        //    "aButtons": ["csv", "xls", "pdf"]
                        //}
                    ],
                    "sSwfPath": "assets/swf/copy_csv_xls_pdf.swf"
                },
                "language": {
                    "search": "",
                    "sLengthMenu": "_MENU_",
                    "oPaginate": {
                        "sPrevious": "Prev",
                        "sNext": "Next"
                    }
                },
                "aoColumns": [
                    { "bVisible": false },
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    { "bSortable": false }
                ]
            });

            var isEditing = null;

            //Delete an Existing Row
            $('#editabledatatable').on("click", 'a.delete', function (e) {
                e.preventDefault();
                var nRow = $(this).parents('tr')[0];
                var aData = oTable.fnGetData(nRow);
                var id = aData[0];
                bootbox.confirm("Are you sure to Delete?", function (result) {
                    if (result) {
                        oTable.fnDeleteRow(nRow);
                        DeleteEntityItem(id, "/AppUser/DeleteEntity", "User");
                    }
                });
                return false;
            });

            //Edit A Row
            $('#editabledatatable').on("click", 'a.edit', function (e) {
                e.preventDefault();
                var nRow = $(this).parents('tr')[0];
                var aData = oTable.fnGetData(nRow);
                var entID = aData[0];
                $("#divnew").show();
                $(".topfocus").focus().css('border', '');//To be visible (scroll to top) the edit section by default
                $("#btnUpdate").show();
                $("#btnCreate").hide();
                //Get the data for the ID and bind it to controls
                if (entID != "") {
                    $.ajax({
                        url: "/AppUser/GetDataForMasterID",
                        dataType: "json",
                        type: "POST",
                        beforeSend: ShowProgressBar(),
                        complete: HideProgressBar(),
                        contentType: 'application/json; charset=utf-8',
                        cache: false,
                        data: JSON.stringify({ id: entID }),
                        success: function (data) {
                            retmsg = data.message;
                            if (retmsg == '') {
                                $('.btn-danger').trigger('click');
                                $('.modal-body').html('Error while getting the data.');
                            }
                            else {
                                 //alert(retmsg[1]);
                                $("#divID").html(retmsg.ID);
                                $("#txt1").val(retmsg.FirstName);
                                $("#txt2").val(retmsg.LastName);
                                $("#txt3").val(retmsg.Email);
                                $("#txt4").val(retmsg.Mobile);

                                //var dat = new Date(1501439400000);
                                //alert(dat);
                                //$('#id-date-picker-1').datepicker('setValue', dat);
                                //alert(ToJavaScriptDate(retmsg.DOB));
                                if (retmsg.DOB != null && retmsg.DOB != "")
                                    $('#date1').datepicker('setValue', ToJavaScriptDate(retmsg.DOB));

                                $("#ddlMultiple").multipleSelect("setSelects", data.success);

                                if (retmsg.Gender == "Male")
                                    $("#rbn1").prop("checked", true);
                                else
                                    $("#rbn2").prop("checked", true);

                                if (retmsg.IsActive == true)
                                    $("#chk1").prop('checked', true);//.attr("checked", true);
                                else
                                    $("#chk1").prop('checked', false);//.removeAttr("checked");
                            }
                        },
                        error: function (xhr) {
                            $('.btn-danger').trigger('click');
                            $('.modal-body').html('Error while getting the data.');
                        }
                    });
                }

            });
        }
    };
}();


var InitiateEditableDataTable_Test = function () {
    return {
        init: function () {
            //Datatable Initiating
            var oTable = $('#editabledatatable').dataTable({
                "aLengthMenu": [
                    [5, 10, 20, 100, -1],
                    [5, 10, 20, 100, "All"]
                ],
                "iDisplayLength": 10,
                "sPaginationType": "bootstrap",
                "sDom": "Tflt<'row DTTTFooter'<'col-sm-6'i><'col-sm-6'p>>",
                "oTableTools": {
                    "aButtons": [
                        //"copy",
                        //"print",                           
                        //{
                        //    "sExtends": "collection",
                        //    "sButtonText": "Save <i class=\"fa fa-angle-down\"></i>",
                        //    "aButtons": ["csv", "xls", "pdf"]
                        //}
                    ],
                    "sSwfPath": "assets/swf/copy_csv_xls_pdf.swf"
                },
                "language": {
                    "search": "",
                    "sLengthMenu": "_MENU_",
                    "oPaginate": {
                        "sPrevious": "Prev",
                        "sNext": "Next"
                    }
                },
                "aoColumns": [
                    { "bVisible": false },
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    { "bSortable": false }
                ]
            });

            var isEditing = null;

            //Delete an Existing Row
            $('#editabledatatable').on("click", 'a.delete', function (e) {
                e.preventDefault();
                var nRow = $(this).parents('tr')[0];
                var aData = oTable.fnGetData(nRow);
                var id = aData[0];
                bootbox.confirm("Are you sure to Delete?", function (result) {
                    if (result) {
                        oTable.fnDeleteRow(nRow);
                        DeleteEntityItem(id, "/AppUser/DeleteEntity", "User");
                    }
                });
                return false;
            });

            //Edit A Row
            $('#editabledatatable').on("click", 'a.edit', function (e) {
                e.preventDefault();
                var nRow = $(this).parents('tr')[0];
                var aData = oTable.fnGetData(nRow);
                var entID = aData[0];
                $("#divnew").show();
                $(".topfocus").focus().css('border', '');//To be visible (scroll to top) the edit section by default
                $("#btnUpdate").show();
                $("#btnCreate").hide();
                //Get the data for the ID and bind it to controls
                if (entID != "") {
                    $.ajax({
                        url: "/AppUser/GetDataForMasterID",
                        dataType: "json",
                        type: "POST",
                        beforeSend: ShowProgressBar(),
                        complete: HideProgressBar(),
                        contentType: 'application/json; charset=utf-8',
                        cache: false,
                        data: JSON.stringify({ id: entID }),
                        success: function (data) {
                            retmsg = data.message;
                            if (retmsg == '') {
                                $('.btn-danger').trigger('click');
                                $('.modal-body').html('Error while getting the data.');
                            }
                            else {
                                //alert(retmsg[1]);
                                $("#divID").html(retmsg.ID);
                                $("#txt1").val(retmsg.FirstName);
                                $("#txt2").val(retmsg.LastName);
                                $("#txt3").val(retmsg.Email);
                                $("#txt4").val(retmsg.Mobile);

                                //var dat = new Date(1501439400000);
                                //alert(dat);
                                //$('#id-date-picker-1').datepicker('setValue', dat);
                                //alert(ToJavaScriptDate(retmsg.DOB));
                                if (retmsg.DOB != null && retmsg.DOB != "")
                                    $('#date1').datepicker('setValue', ToJavaScriptDate(retmsg.DOB));

                                $("#ddlMultiple").multipleSelect("setSelects", data.success);

                                if (retmsg.Gender == "Male")
                                    $("#rbn1").prop("checked", true);
                                else
                                    $("#rbn2").prop("checked", true);

                                if (retmsg.IsActive == true)
                                    $("#chk1").prop('checked', true);//.attr("checked", true);
                                else
                                    $("#chk1").prop('checked', false);//.removeAttr("checked");
                            }
                        },
                        error: function (xhr) {
                            $('.btn-danger').trigger('click');
                            $('.modal-body').html('Error while getting the data.');
                        }
                    });
                }

            });
        }
    };
}();

var InitiateInboxDataTable = function () {
    return {
        init: function () {
            //Datatable Initiating
            var oTable = $('#editabledatatable').dataTable({
                "aLengthMenu": [
                    [5, 10, 20, 100, -1],
                    [5, 10, 20, 100, "All"]
                ],
                "iDisplayLength": 10,
                "sPaginationType": "bootstrap",
                "sDom": "Tflt<'row DTTTFooter'<'col-sm-6'i><'col-sm-6'p>>",
                "oTableTools": {
                    "aButtons": [
                        //"copy",
                        //"print",                           
                        //{
                        //    "sExtends": "collection",
                        //    "sButtonText": "Save <i class=\"fa fa-angle-down\"></i>",
                        //    "aButtons": ["csv", "xls", "pdf"]
                        //}
                    ],
                    "sSwfPath": "assets/swf/copy_csv_xls_pdf.swf"
                },
                "language": {
                    "search": "",
                    "sLengthMenu": "_MENU_",
                    "oPaginate": {
                        "sPrevious": "Prev",
                        "sNext": "Next"
                    }
                },
                "aoColumns": [
                    { "bVisible": false },
                    { "bVisible": false },
                    null,
                    null,                    
                    null,
                    null,
                    null,
                    null,
                    null,
                    { "bSortable": false }
                ]
            });

            var isEditing = null;

            //Delete an Existing Row
            $('#editabledatatable').on("click", 'a.delete', function (e) {
                e.preventDefault();
                var nRow = $(this).parents('tr')[0];
                var aData = oTable.fnGetData(nRow);
                var id = aData[0];
                bootbox.confirm("Are you sure to Delete?", function (result) {
                    if (result) {
                        oTable.fnDeleteRow(nRow);
                        DeleteEntityItem(id, "/AppUser/DeleteEntity", "User");
                    }
                });
                return false;
            });

            //Edit A Row
            $('#editabledatatable').on("click", 'a.edit', function (e) {
                e.preventDefault();
                var nRow = $(this).parents('tr')[0];
                var aData = oTable.fnGetData(nRow);
                var entID = aData[0];
                $("#divnew").show();
                $(".topfocus").focus().css('border', '');//To be visible (scroll to top) the edit section by default
                $("#btnUpdate").show();
                $("#btnCreate").hide();
                //Get the data for the ID and bind it to controls
                if (entID != "") {
                    $.ajax({
                        url: "/AppUser/GetDataForMasterID",
                        dataType: "json",
                        type: "POST",
                        beforeSend: ShowProgressBar(),
                        complete: HideProgressBar(),
                        contentType: 'application/json; charset=utf-8',
                        cache: false,
                        data: JSON.stringify({ id: entID }),
                        success: function (data) {
                            retmsg = data.message;
                            if (retmsg == '') {
                                $('.btn-danger').trigger('click');
                                $('.modal-body').html('Error while getting the data.');
                            }
                            else {
                                //alert(retmsg[1]);
                                $("#divID").html(retmsg.ID);
                                $("#txt1").val(retmsg.FirstName);
                                $("#txt2").val(retmsg.LastName);
                                $("#txt3").val(retmsg.Email);
                                $("#txt4").val(retmsg.Mobile);

                                //var dat = new Date(1501439400000);
                                //alert(dat);
                                //$('#id-date-picker-1').datepicker('setValue', dat);
                                //alert(ToJavaScriptDate(retmsg.DOB));
                                if (retmsg.DOB != null && retmsg.DOB != "")
                                    $('#date1').datepicker('setValue', ToJavaScriptDate(retmsg.DOB));

                                $("#ddlMultiple").multipleSelect("setSelects", data.success);

                                if (retmsg.Gender == "Male")
                                    $("#rbn1").prop("checked", true);
                                else
                                    $("#rbn2").prop("checked", true);

                                if (retmsg.IsActive == true)
                                    $("#chk1").prop('checked', true);//.attr("checked", true);
                                else
                                    $("#chk1").prop('checked', false);//.removeAttr("checked");
                            }
                        },
                        error: function (xhr) {
                            $('.btn-danger').trigger('click');
                            $('.modal-body').html('Error while getting the data.');
                        }
                    });
                }

            });
        }
    };
}();


var InitiateApproverInboxDataTable = function () {
    return {
        init: function () {
            //Datatable Initiating
            var oTable = $('#editabledatatableApprover').dataTable({
                "aLengthMenu": [
                    [5, 10, 20, 100, -1],
                    [5, 10, 20, 100, "All"]
                ],
                "iDisplayLength": 10,
                "sPaginationType": "bootstrap",
                "sDom": "Tflt<'row DTTTFooter'<'col-sm-6'i><'col-sm-6'p>>",
                "oTableTools": {
                    "aButtons": [
                        //"copy",
                        //"print",                           
                        //{
                        //    "sExtends": "collection",
                        //    "sButtonText": "Save <i class=\"fa fa-angle-down\"></i>",
                        //    "aButtons": ["csv", "xls", "pdf"]
                        //}
                    ],
                    "sSwfPath": "assets/swf/copy_csv_xls_pdf.swf"
                },
                "language": {
                    "search": "",
                    "sLengthMenu": "_MENU_",
                    "oPaginate": {
                        "sPrevious": "Prev",
                        "sNext": "Next"
                    }
                },
                "aoColumns": [
                    { "bVisible": false },
                    { "bVisible": false },
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    { "bSortable": false }
                ]
            });

            var isEditing = null;

            //Delete an Existing Row
            $('#editabledatatableApprover').on("click", 'a.delete', function (e) {
                e.preventDefault();
                var nRow = $(this).parents('tr')[0];
                var aData = oTable.fnGetData(nRow);
                var id = aData[0];
                bootbox.confirm("Are you sure to Delete?", function (result) {
                    if (result) {
                        oTable.fnDeleteRow(nRow);
                        DeleteEntityItem(id, "/AppUser/DeleteEntity", "User");
                    }
                });
                return false;
            });

            //Edit A Row
            $('#editabledatatableApprover').on("click", 'a.edit', function (e) {
                e.preventDefault();
                var nRow = $(this).parents('tr')[0];
                var aData = oTable.fnGetData(nRow);
                var entID = aData[0];
                $("#divnew").show();
                $(".topfocus").focus().css('border', '');//To be visible (scroll to top) the edit section by default
                $("#btnUpdate").show();
                $("#btnCreate").hide();
                //Get the data for the ID and bind it to controls
                if (entID != "") {
                    $.ajax({
                        url: "/AppUser/GetDataForMasterID",
                        dataType: "json",
                        type: "POST",
                        beforeSend: ShowProgressBar(),
                        complete: HideProgressBar(),
                        contentType: 'application/json; charset=utf-8',
                        cache: false,
                        data: JSON.stringify({ id: entID }),
                        success: function (data) {
                            retmsg = data.message;
                            if (retmsg == '') {
                                $('.btn-danger').trigger('click');
                                $('.modal-body').html('Error while getting the data.');
                            }
                            else {
                                //alert(retmsg[1]);
                                $("#divID").html(retmsg.ID);
                                $("#txt1").val(retmsg.FirstName);
                                $("#txt2").val(retmsg.LastName);
                                $("#txt3").val(retmsg.Email);
                                $("#txt4").val(retmsg.Mobile);

                                //var dat = new Date(1501439400000);
                                //alert(dat);
                                //$('#id-date-picker-1').datepicker('setValue', dat);
                                //alert(ToJavaScriptDate(retmsg.DOB));
                                if (retmsg.DOB != null && retmsg.DOB != "")
                                    $('#date1').datepicker('setValue', ToJavaScriptDate(retmsg.DOB));

                                $("#ddlMultiple").multipleSelect("setSelects", data.success);

                                if (retmsg.Gender == "Male")
                                    $("#rbn1").prop("checked", true);
                                else
                                    $("#rbn2").prop("checked", true);

                                if (retmsg.IsActive == true)
                                    $("#chk1").prop('checked', true);//.attr("checked", true);
                                else
                                    $("#chk1").prop('checked', false);//.removeAttr("checked");
                            }
                        },
                        error: function (xhr) {
                            $('.btn-danger').trigger('click');
                            $('.modal-body').html('Error while getting the data.');
                        }
                    });
                }

            });
        }
    };
}();

var InitiateQMSInboxDataTable = function () {
    return {
        init: function () {
            //Datatable Initiating
            var oTable = $('#editabledatatable_QMS').dataTable({
                "aLengthMenu": [
                    [5, 10, 20, 100, -1],
                    [5, 10, 20, 100, "All"]
                ],
                "iDisplayLength": 10,
                "sPaginationType": "bootstrap",
                "sDom": "Tflt<'row DTTTFooter'<'col-sm-6'i><'col-sm-6'p>>",
                "oTableTools": {
                    "aButtons": [
                        //"copy",
                        //"print",                           
                        //{
                        //    "sExtends": "collection",
                        //    "sButtonText": "Save <i class=\"fa fa-angle-down\"></i>",
                        //    "aButtons": ["csv", "xls", "pdf"]
                        //}
                    ],
                    "sSwfPath": "assets/swf/copy_csv_xls_pdf.swf"
                },
                "language": {
                    "search": "",
                    "sLengthMenu": "_MENU_",
                    "oPaginate": {
                        "sPrevious": "Prev",
                        "sNext": "Next"
                    }
                },
                "aoColumns": [
                    { "bVisible": false },
                    { "bVisible": false },
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    { "bSortable": false }
                ]
            });

        }
    };
}();


var InitiateWFHistoryDataTable = function () {
    return {
        init: function () {
            //Datatable Initiating
            var oTable = $('#editabledatatable').dataTable({
                "aLengthMenu": [
                    [5, 10, 20, 100, -1],
                    [5, 10, 20, 100, "All"]
                ],
                "iDisplayLength": 20,
                "sPaginationType": "bootstrap",
                "sDom": "Tflt<'row DTTTFooter'<'col-sm-6'i><'col-sm-6'p>>",
                "oTableTools": {
                    "aButtons": [
                        //"copy",
                        //"print",                           
                        //{
                        //    "sExtends": "collection",
                        //    "sButtonText": "Save <i class=\"fa fa-angle-down\"></i>",
                        //    "aButtons": ["csv", "xls", "pdf"]
                        //}
                    ],
                    "sSwfPath": "assets/swf/copy_csv_xls_pdf.swf"
                },
                "language": {
                    "search": "",
                    "sLengthMenu": "_MENU_",
                    "oPaginate": {
                        "sPrevious": "Prev",
                        "sNext": "Next"
                    }
                },
                "aoColumns": [
                    { "bVisible": false },
                    { "bVisible": false },
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    { "bSortable": false }
                ]
            });

            var isEditing = null;

            //Delete an Existing Row
            $('#editabledatatable').on("click", 'a.delete', function (e) {
                e.preventDefault();
                var nRow = $(this).parents('tr')[0];
                var aData = oTable.fnGetData(nRow);
                var id = aData[0];
                bootbox.confirm("Are you sure to Delete?", function (result) {
                    if (result) {
                        oTable.fnDeleteRow(nRow);
                        DeleteEntityItem(id, "/AppUser/DeleteEntity", "User");
                    }
                });
                return false;
            });

            //Edit A Row
            $('#editabledatatable').on("click", 'a.edit', function (e) {
                e.preventDefault();
                var nRow = $(this).parents('tr')[0];
                var aData = oTable.fnGetData(nRow);
                var entID = aData[0];
                $("#divnew").show();
                $(".topfocus").focus().css('border', '');//To be visible (scroll to top) the edit section by default
                $("#btnUpdate").show();
                $("#btnCreate").hide();
                //Get the data for the ID and bind it to controls
                if (entID != "") {
                    $.ajax({
                        url: "/AppUser/GetDataForMasterID",
                        dataType: "json",
                        type: "POST",
                        beforeSend: ShowProgressBar(),
                        complete: HideProgressBar(),
                        contentType: 'application/json; charset=utf-8',
                        cache: false,
                        data: JSON.stringify({ id: entID }),
                        success: function (data) {
                            retmsg = data.message;
                            if (retmsg == '') {
                                $('.btn-danger').trigger('click');
                                $('.modal-body').html('Error while getting the data.');
                            }
                            else {
                                //alert(retmsg[1]);
                                $("#divID").html(retmsg.ID);
                                $("#txt1").val(retmsg.FirstName);
                                $("#txt2").val(retmsg.LastName);
                                $("#txt3").val(retmsg.Email);
                                $("#txt4").val(retmsg.Mobile);

                                //var dat = new Date(1501439400000);
                                //alert(dat);
                                //$('#id-date-picker-1').datepicker('setValue', dat);
                                //alert(ToJavaScriptDate(retmsg.DOB));
                                if (retmsg.DOB != null && retmsg.DOB != "")
                                    $('#date1').datepicker('setValue', ToJavaScriptDate(retmsg.DOB));

                                $("#ddlMultiple").multipleSelect("setSelects", data.success);

                                if (retmsg.Gender == "Male")
                                    $("#rbn1").prop("checked", true);
                                else
                                    $("#rbn2").prop("checked", true);

                                if (retmsg.IsActive == true)
                                    $("#chk1").prop('checked', true);//.attr("checked", true);
                                else
                                    $("#chk1").prop('checked', false);//.removeAttr("checked");
                            }
                        },
                        error: function (xhr) {
                            $('.btn-danger').trigger('click');
                            $('.modal-body').html('Error while getting the data.');
                        }
                    });
                }

            });
        }
    };
}();


var InitiateEditableUserDataTable = function () {
    return {
        init: function () {
            //Datatable Initiating
            var oTable = $('#editabledatatable').dataTable({
                "aLengthMenu": [
                    [5, 10, 20, 100, -1],
                    [5, 10, 20, 100, "All"]
                ],
                "iDisplayLength": 10,
                "sPaginationType": "bootstrap",
                "sDom": "Tflt<'row DTTTFooter'<'col-sm-6'i><'col-sm-6'p>>",
                "oTableTools": {
                    "aButtons": [
                        ////"copy",
                        ////"print",
                        //{
                        //    "sExtends": "collection",
                        //    "sButtonText": "Save <i class=\"fa fa-angle-down\"></i>",
                        //    "aButtons": ["csv", "xls", "pdf"]
                        //}
                    ],
                    "sSwfPath": "assets/swf/copy_csv_xls_pdf.swf"
                },
                "language": {
                    "search": "",
                    "sLengthMenu": "_MENU_",
                    "oPaginate": {
                        "sPrevious": "Prev",
                        "sNext": "Next"
                    }
                },
                "aoColumns": [
                    null,
                    null,
                    null,
                    null,
                    { "bSortable": false }
                ]
            });

            var isEditing = null;

            //Add New Row
            $('#editabledatatable_new').click(function (e) {
                e.preventDefault();
                var aiNew = oTable.fnAddData([
                    '', '', '', '',
                    '<a href="#" class="btn btn-success btn-xs save"><i class="fa fa-edit"></i> Save</a> <a href="#" class="btn btn-warning btn-xs cancel" data-mode="new"><i class="fa fa-times"></i> Cancel</a>'
                ]);
                var nRow = oTable.fnGetNodes(aiNew[0]);
                editAddedRow(oTable, nRow);
                isEditing = nRow;
            });

            //Delete an Existing Row
            $('#editabledatatable').on("click", 'a.delete', function (e) {
                e.preventDefault();

                var nRow = $(this).parents('tr')[0];
                var aData = oTable.fnGetData(nRow);
                var id = aData[0];

                bootbox.confirm("Are you sure to Delete?", function (result) {
                    if (result) {
                        oTable.fnDeleteRow(nRow);
                        DeleteUser(id);
                    }
                });
            });

            //Cancel Editing or Adding a Row
            $('#editabledatatable').on("click", 'a.cancel', function (e) {
                e.preventDefault();
                if ($(this).attr("data-mode") == "new") {
                    var nRow = $(this).parents('tr')[0];
                    oTable.fnDeleteRow(nRow);
                    isEditing = null;
                } else {
                    restoreRow(oTable, isEditing);
                    isEditing = null;
                }
            });

            //Edit A Row
            $('#editabledatatable').on("click", 'a.edit', function (e) {
                e.preventDefault();
                var nRow = $(this).parents('tr')[0];
                //alert(nRow); alert(isEditing);

                if (isEditing !== null && isEditing != nRow) {
                    restoreRow(oTable, isEditing);
                    editRow(oTable, nRow);
                    isEditing = nRow;
                } else {
                    editRow(oTable, nRow);
                    isEditing = nRow;
                }
            });

            //Save an Editing Row
            $('#editabledatatable').on("click", 'a.save', function (e) {
                e.preventDefault();

                if (this.innerHTML.indexOf("Save") >= 0) {
                    var thisRow = isEditing;
                    var jqInputs = $('input', isEditing);
                    //alert('save row');
                    //New row adding
                    if (jqInputs[0].value.trim() == '') {
                        if (checkForUserSave(jqInputs[1].value, jqInputs[2].value, jqInputs[3].value)) {
                            saveRow(oTable, isEditing, 'new');
                            isEditing = null;
                        }
                        else {
                            var aiNew = oTable.fnAddData([
                        '', jqInputs[1].value, jqInputs[2].value, jqInputs[3].value,
                        '<a href="#" class="btn btn-success btn-xs save"><i class="fa fa-edit"></i> Save</a> <a href="#" class="btn btn-warning btn-xs cancel" data-mode="new"><i class="fa fa-times"></i> Cancel</a>'
                            ]);
                            var nRow = oTable.fnGetNodes(aiNew[0]);
                            editAddedRowWithBorder(oTable, nRow);
                            isEditing = nRow;
                        }
                    }
                    else {
                        var isvalid = checkForUserSave(jqInputs[1].value, jqInputs[2].value, jqInputs[3].value);

                        if (isvalid) {
                            saveRow(oTable, isEditing, 'edit');
                            isEditing = null;
                        }
                        else {
                            isEditing = thisRow;

                            var jqTds = $('>td', thisRow);
                            jqTds[0].innerHTML = '<input style="width:50px;" readonly="readonly" type="text" class="form-control input-small" value="' + jqInputs[0].value + '">';
                            if (jqInputs[1].value.trim() == '')
                                jqTds[1].innerHTML = '<input style="border:red 1px solid" maxlength="50" type="text" class="form-control input-small" value="' + jqInputs[1].value + '">';
                            else
                                jqTds[1].innerHTML = '<input maxlength="50" type="text" class="form-control input-small" value="' + jqInputs[1].value + '">';
                            if (jqInputs[2].value.trim() == '')
                                jqTds[2].innerHTML = '<input style="border:red 1px solid" id="txtPhone" type="text" class="form-control input-small" value="' + jqInputs[2].value + '">';
                            else {
                                if (!ValidateEmail(jqInputs[2].value.trim())) {
                                    jqTds[2].innerHTML = '<input style="border:red 1px solid" id="txtPhone" type="text" class="form-control input-small" value="' + jqInputs[2].value + '">';
                                }
                                else
                                    jqTds[2].innerHTML = '<input id="txtPhone" type="text" class="form-control input-small" value="' + jqInputs[2].value + '">';
                            }
                            if (jqInputs[3].value.trim() == '')
                                jqTds[3].innerHTML = '<input style="border:red 1px solid" onkeydown="Allownumbers(event);" maxlength="10"  type="text" class="form-control input-small" value="' + jqInputs[3].value + '">';
                            else
                                jqTds[3].innerHTML = '<input onkeydown="Allownumbers(event);" maxlength="10"  type="text" class="form-control input-small" value="' + jqInputs[3].value + '">';
                            jqTds[4].innerHTML = '<a href="#" class="btn btn-success btn-xs save"><i class="fa fa-save"></i> Save</a> <a href="#" class="btn btn-warning btn-xs cancel"><i class="fa fa-times"></i> Cancel</a>';

                        }
                    }

                    //Some Code to Highlight Updated Row
                }
            });


            function restoreRow(oTable, nRow) {
                //alert('restoreRow');
                var aData = oTable.fnGetData(nRow);
                var jqTds = $('>td', nRow);

                for (var i = 0, iLen = jqTds.length; i < iLen; i++) {
                    oTable.fnUpdate(aData[i], nRow, i, false);
                }

                oTable.fnDraw();
            }

            function editRow(oTable, nRow) {
                //alert('In editRow');
                var aData = oTable.fnGetData(nRow);
                var jqTds = $('>td', nRow);
                //alert(aData[0])
                jqTds[0].innerHTML = '<input style="width:50px;" readonly="readonly" type="text" class="form-control input-small" value="' + aData[0] + '">';
                //jqTds[0].innerHTML = aData[0];
                jqTds[1].innerHTML = '<input maxlength="50" type="text" class="form-control input-small" value="' + aData[1] + '">';
                jqTds[2].innerHTML = '<input id="txtPhone"  type="text" class="form-control input-small" value="' + aData[2] + '">';
                jqTds[3].innerHTML = '<input onkeydown="Allownumbers(event);" maxlength="10" type="text" class="form-control input-small" value="' + aData[3] + '">';
                //jqTds[4].innerHTML = '<input type="text" class="form-control input-small" value="' + aData[4] + '">';
                jqTds[4].innerHTML = '<a href="#" class="btn btn-success btn-xs save"><i class="fa fa-save"></i> Save</a> <a href="#" class="btn btn-warning btn-xs cancel"><i class="fa fa-times"></i> Cancel</a>';
            }

            function editAddedRow(oTable, nRow) {
                //alert('In editAddedRow');
                var aData = oTable.fnGetData(nRow);
                var jqTds = $('>td', nRow);
                jqTds[0].innerHTML = '<input style="width:5px;display:none;" readonly="readonly" type="text" class="form-control input-small" value="' + aData[0] + '">';
                //jqTds[0].innerHTML = aData[0];
                jqTds[1].innerHTML = '<input maxlength="50" type="text" class="form-control input-small" value="' + aData[1] + '">';
                jqTds[2].innerHTML = '<input id="txtPhone"  type="text" class="form-control input-small" value="' + aData[2] + '">';
                jqTds[3].innerHTML = '<input onkeydown="Allownumbers(event);" maxlength="10" type="text" class="form-control input-small" value="' + aData[3] + '">';
                //jqTds[4].innerHTML = '<input type="text" class="form-control input-small" value="' + aData[4] + '">';
                //jqTds[4].innerHTML = '<div class="radio"><lable><input name="form-field-radio" type="radio" class="colored-blue"><span class="text"> Blackberry</span></label></div>';
                jqTds[4].innerHTML = aData[4];

            }

            function editAddedRowWithBorder(oTable, nRow) {
                var aData = oTable.fnGetData(nRow);
                var jqTds = $('>td', nRow);
                jqTds[0].innerHTML = '<input style="width:5px;display:none;" readonly="readonly" type="text" class="form-control input-small" value="' + aData[0] + '">';
                if (aData[1].trim() == '')
                    jqTds[1].innerHTML = '<input style="border:red 1px solid" maxlength="50" type="text" class="form-control input-small" value="' + aData[1] + '">';
                else
                    jqTds[1].innerHTML = '<input maxlength="50" type="text" class="form-control input-small" value="' + aData[1] + '">';
                if (aData[2].trim() == '')
                    jqTds[2].innerHTML = '<input style="border:red 1px solid" id="txtPhone"  type="text" class="form-control input-small" value="' + aData[2] + '">';
                else {
                    if (!ValidateEmail(aData[2].trim())) {
                        jqTds[2].innerHTML = '<input style="border:red 1px solid" id="txtPhone"  type="text" class="form-control input-small" value="' + aData[2] + '">';
                    }
                    else
                        jqTds[2].innerHTML = '<input id="txtPhone"  type="text" class="form-control input-small" value="' + aData[2] + '">';
                }
                if (aData[3].trim() == '')
                    jqTds[3].innerHTML = '<input style="border:red 1px solid" onkeydown="Allownumbers(event);" maxlength="10" type="text" class="form-control input-small" value="' + aData[3] + '">';
                else {                    
                    jqTds[3].innerHTML = '<input onkeydown="Allownumbers(event);" maxlength="10" type="text" class="form-control input-small" value="' + aData[3] + '">';
                }
                jqTds[4].innerHTML = aData[4];
            }

            function saveRow(oTable, nRow, action) {
                //alert('save row');
                var jqInputs = $('input', nRow);
                //alert(jqInputs[4].value);
                oTable.fnUpdate(jqInputs[0].value, nRow, 0, false);
                oTable.fnUpdate(jqInputs[1].value, nRow, 1, false);
                oTable.fnUpdate(jqInputs[2].value, nRow, 2, false);
                oTable.fnUpdate(jqInputs[3].value, nRow, 3, false);
                //oTable.fnUpdate(jqInputs[4].value, nRow, 4, false);
                oTable.fnUpdate('<a href="#" class="btn btn-azure btn-xs edit"><i class="fa fa-edit"></i> Edit</a> <a href="#" class="btn btn-redcolor btn-xs delete"><i class="fa fa-trash-o"></i> Delete</a>', nRow, 4, false);
                var msg = SaveUserToDB(jqInputs[0].value, jqInputs[1].value, jqInputs[2].value, jqInputs[3].value);
                if (action == 'new') {
                    var arr = msg.split('#');
                    oTable.fnUpdate(arr[0], nRow, 0, false);
                }
                if (action == 'edit' && msg.indexOf("#") >= 0) {
                    var arrst = msg.split('#');
                    //alert(arrst);
                    oTable.fnUpdate(arrst[0], nRow, 1, false);
                    oTable.fnUpdate(arrst[1], nRow, 2, false);
                    //oTable.fnUpdate(arrst[4], nRow, 3, false);
                    oTable.fnUpdate(arrst[2], nRow, 3, false);
                }
                oTable.fnDraw();
                if (msg == '0' && action == 'new') {
                    //alert('Already Exists as');
                    oTable.fnDeleteRow(nRow);
                }
            }
            function cancelEditRow(oTable, nRow) {
                var jqInputs = $('input', nRow);
                oTable.fnUpdate(jqInputs[0].value, nRow, 0, false);
                oTable.fnUpdate(jqInputs[1].value, nRow, 1, false);
                oTable.fnUpdate(jqInputs[2].value, nRow, 2, false);
                oTable.fnUpdate(jqInputs[3].value, nRow, 3, false);
                //oTable.fnUpdate(jqInputs[4].value, nRow, 4, false);
                oTable.fnUpdate('<a href="#" class="btn btn-azure btn-xs edit"><i class="fa fa-edit"></i> Edit</a> <a href="#" class="btn btn-redcolor btn-xs delete"><i class="fa fa-trash-o"></i> Delete</a>', nRow, 4, false);
                oTable.fnDraw();
            }
        }

    };
}();

var InitiateEditableAppUserUserDataTable1 = function () {
    return {
        init: function () {
            //Datatable Initiating
            var oTable = $('#editabledatatable').dataTable({
                "aLengthMenu": [
                    [5, 10, 20, 100, -1],
                    [5, 10, 20, 100, "All"]
                ],
                "iDisplayLength": 10,
                "sPaginationType": "bootstrap",
                "sDom": "Tflt<'row DTTTFooter'<'col-sm-6'i><'col-sm-6'p>>",
                "oTableTools": {
                    "aButtons": [
                        ////"copy",
                        ////"print",                           
                        //{
                        //    "sExtends": "collection",
                        //    "sButtonText": "Save <i class=\"fa fa-angle-down\"></i>",
                        //    "aButtons": ["csv", "xls", "pdf"]
                        //}
                    ],
                    "sSwfPath": "assets/swf/copy_csv_xls_pdf.swf"
                },
                "language": {
                    "search": "",
                    "sLengthMenu": "_MENU_",
                    "oPaginate": {
                        "sPrevious": "Prev",
                        "sNext": "Next"
                    }
                },
                "aoColumns": [
                    { "bVisible": false },
                    null,
                    null,
                    null,
                    null,
                    { "bSortable": false }
                ]
            });

            var isEditing = null;

            //Delete an Existing Row
            $('#editabledatatable').on("click", 'a.delete', function (e) {
                e.preventDefault();
                var nRow = $(this).parents('tr')[0];
                var aData = oTable.fnGetData(nRow);
                var id = aData[0];
                bootbox.confirm("Are you sure to Delete?", function (result) {
                    if (result) {
                        oTable.fnDeleteRow(nRow);                        
                        DeleteEntityItem(id, "/AppUser/DeleteEntity", "User");
                    }
                });
                return false;
            });
           
            //Edit A Row
            $('#editabledatatable').on("click", 'a.edit', function (e) {
                e.preventDefault();
                var nRow = $(this).parents('tr')[0];
                var aData = oTable.fnGetData(nRow);
                var entID = aData[0];
                $("#divnew").show();
                $(".topfocus").focus().css('border','');//To be visible (scroll to top) the edit section by default
                $("#btnUpdate").show();
                $("#btnCreate").hide();
                //Get the data for the ID and bind it to controls
                if(entID !="")
                {
                    $.ajax({
                        url: "/AppUser/GetDataForMasterID",
                        dataType: "json",
                        type: "POST",
                        beforeSend: ShowProgressBar(),
                        complete: HideProgressBar(),
                        contentType: 'application/json; charset=utf-8',
                        cache: false,
                        data: JSON.stringify({ id: entID }),
                        success: function (data) {
                            retmsg = data.message;
                            if (retmsg == '') {
                                $('.btn-danger').trigger('click');
                                $('.modal-body').html('Error while getting the data.');
                            }
                            else {
                                //alert(retmsg[1]);
                                $("#divID").html(retmsg.ID);
                                $("#txt1").val(retmsg.FirstName);
                                $("#txt2").val(retmsg.LastName);
                                $("#txt3").val(retmsg.Email);
                                $("#txt4").val(retmsg.Mobile);

                                //var dat = new Date(1501439400000);
                                //alert(dat);
                                //$('#id-date-picker-1').datepicker('setValue', dat);
                                //alert(ToJavaScriptDate(retmsg.DOB));
                                if (retmsg.DOB != null && retmsg.DOB != "")
                                    $('#date1').datepicker('setValue', ToJavaScriptDate(retmsg.DOB));

                                $("#ddlMultiple").multipleSelect("setSelects", data.success);
                                
                                if (retmsg.Gender == "Male")
                                    $("#rbn1").prop("checked", true);
                                else
                                    $("#rbn2").prop("checked", true);
                                
                                if (retmsg.IsActive == true)
                                    $("#chk1").prop('checked', true);//.attr("checked", true);
                                else
                                    $("#chk1").prop('checked', false);//.removeAttr("checked");
                            }
                        },
                        error: function (xhr) {
                            $('.btn-danger').trigger('click');
                            $('.modal-body').html('Error while getting the data.');
                        }
                    });
                }

            });
        }
    };
}();

var InitiatePublishedDataTable = function () {
    return {
        init: function () {
            //Datatable Initiating
            var oTable = $('#editabledatatable').dataTable({
                "aLengthMenu": [
                    [5, 10, 20, 100, -1],
                    [5, 10, 20, 100, "All"]
                ],
                "iDisplayLength": 10,
                "sPaginationType": "bootstrap",
                "sDom": "Tflt<'row DTTTFooter'<'col-sm-6'i><'col-sm-6'p>>",
                "oTableTools": {
                    "aButtons": [
                        //"copy",
                        //"print",                           
                        //{
                        //    "sExtends": "collection",
                        //    "sButtonText": "Save <i class=\"fa fa-angle-down\"></i>",
                        //    "aButtons": ["csv", "xls", "pdf"]
                        //}
                    ],
                    "sSwfPath": "assets/swf/copy_csv_xls_pdf.swf"
                },
                "language": {
                    "search": "",
                    "sLengthMenu": "_MENU_",
                    "oPaginate": {
                        "sPrevious": "Prev",
                        "sNext": "Next"
                    }
                },
                "aoColumns": [
                    { "bVisible": false },
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null
                ]
            });
        }
    };
}();

var InitiatePendingDataTable = function () {
    return {
        init: function () {
            //Datatable Initiating
            var oTable = $('#editabledatatable_1').dataTable({
                "aLengthMenu": [
                    [5, 10, 20, 100, -1],
                    [5, 10, 20, 100, "All"]
                ],
                "iDisplayLength": 10,
                "sPaginationType": "bootstrap",
                "sDom": "Tflt<'row DTTTFooter'<'col-sm-6'i><'col-sm-6'p>>",
                "oTableTools": {
                    "aButtons": [
                        //"copy",
                        //"print",                           
                        //{
                        //    "sExtends": "collection",
                        //    "sButtonText": "Save <i class=\"fa fa-angle-down\"></i>",
                        //    "aButtons": ["csv", "xls", "pdf"]
                        //}
                    ],
                    "sSwfPath": "assets/swf/copy_csv_xls_pdf.swf"
                },
                "language": {
                    "search": "",
                    "sLengthMenu": "_MENU_",
                    "oPaginate": {
                        "sPrevious": "Prev",
                        "sNext": "Next"
                    }
                },
                "aoColumns": [
                    { "bVisible": false },
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null
                ]
            });
        }
    };
}();

var InitiateArchivalProjectDataTable = function () {
    return {
        init: function () {
            //Datatable Initiating
            var oTable = $('#editabledatatable').dataTable({
                "aLengthMenu": [
                    [5, 10, 20, 100, -1],
                    [5, 10, 20, 100, "All"]
                ],
                "iDisplayLength": 10,
                "sPaginationType": "bootstrap",
                "sDom": "Tflt<'row DTTTFooter'<'col-sm-6'i><'col-sm-6'p>>",
                "oTableTools": {
                    "aButtons": [
                        //"copy",
                        //"print",                           
                        //{
                        //    "sExtends": "collection",
                        //    "sButtonText": "Save <i class=\"fa fa-angle-down\"></i>",
                        //    "aButtons": ["csv", "xls", "pdf"]
                        //}
                    ],
                    "sSwfPath": "assets/swf/copy_csv_xls_pdf.swf"
                },
                "language": {
                    "search": "",
                    "sLengthMenu": "_MENU_",
                    "oPaginate": {
                        "sPrevious": "Prev",
                        "sNext": "Next"
                    }
                },
                "aoColumns": [
                    { "bVisible": false },
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null
                ]
            });
        }
    };
}();

var InitiateTemplateDataTable = function () {
    return {
        init: function () {
            //Datatable Initiating
            var oTable = $('#editabledatatable').dataTable({
                "aLengthMenu": [
                    [5, 10, 20, 100, -1],
                    [5, 10, 20, 100, "All"]
                ],
                "iDisplayLength": 20,
                "sPaginationType": "bootstrap",
                "sDom": "Tflt<'row DTTTFooter'<'col-sm-6'i><'col-sm-6'p>>",
                "oTableTools": {
                    "aButtons": [
                        //"copy",
                        //"print",                           
                        //{
                        //    "sExtends": "collection",
                        //    "sButtonText": "Save <i class=\"fa fa-angle-down\"></i>",
                        //    "aButtons": ["csv", "xls", "pdf"]
                        //}
                    ],
                    "sSwfPath": "assets/swf/copy_csv_xls_pdf.swf"
                },
                "language": {
                    "search": "",
                    "sLengthMenu": "_MENU_",
                    "oPaginate": {
                        "sPrevious": "Prev",
                        "sNext": "Next"
                    }
                },
                "aoColumns": [
                    { "bVisible": false },
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    { "bSortable": false },
                ]
            });
        }
    };
}();

var InitiateTemplateDownloadDataTable = function () {
    return {
        init: function () {
            //Datatable Initiating
            var oTable = $('#editabledatatable').dataTable({
                "aLengthMenu": [
                    [5, 10, 20, 100, -1],
                    [5, 10, 20, 100, "All"]
                ],
                "iDisplayLength": 20,
                "sPaginationType": "bootstrap",
                "sDom": "Tflt<'row DTTTFooter'<'col-sm-6'i><'col-sm-6'p>>",
                "oTableTools": {
                    "aButtons": [
                        //"copy",
                        //"print",                           
                        //{
                        //    "sExtends": "collection",
                        //    "sButtonText": "Save <i class=\"fa fa-angle-down\"></i>",
                        //    "aButtons": ["csv", "xls", "pdf"]
                        //}
                    ],
                    "sSwfPath": "assets/swf/copy_csv_xls_pdf.swf"
                },
                "language": {
                    "search": "",
                    "sLengthMenu": "_MENU_",
                    "oPaginate": {
                        "sPrevious": "Prev",
                        "sNext": "Next"
                    }
                },
                "aoColumns": [
                    { "bVisible": false },
                    null,
                    null,
                    null,
                    null,
                    null
                ]
            });
        }
    };
}();

var InitiateDocumentNoDataTable = function () {
    return {
        init: function () {
            //Datatable Initiating
            var oTable = $('#editabledatatable').dataTable({
                "aLengthMenu": [
                    [5, 10, 20, 100, -1],
                    [5, 10, 20, 100, "All"]
                ],
                "iDisplayLength": 10,
                "sPaginationType": "bootstrap",
                "sDom": "Tflt<'row DTTTFooter'<'col-sm-6'i><'col-sm-6'p>>",
                "oTableTools": {
                    "aButtons": [
                        //"copy",
                        //"print",                           
                        //{
                        //    "sExtends": "collection",
                        //    "sButtonText": "Save <i class=\"fa fa-angle-down\"></i>",
                        //    "aButtons": ["csv", "xls", "pdf"]
                        //}
                    ],
                    "sSwfPath": "assets/swf/copy_csv_xls_pdf.swf"
                },
                "language": {
                    "search": "",
                    "sLengthMenu": "_MENU_",
                    "oPaginate": {
                        "sPrevious": "Prev",
                        "sNext": "Next"
                    }
                },
                "aoColumns": [
                    { "bVisible": false },
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null
                ]
            });
        }
    };
}();

var InitiateDepartmentDataTable = function () {
    return {
        init: function () {
            //Datatable Initiating
            var oTable = $('#editabledatatable').dataTable({
                "aLengthMenu": [
                    [5, 10, 20, 100, -1],
                    [5, 10, 20, 100, "All"]
                ],
                "iDisplayLength": 10,
                "sPaginationType": "bootstrap",
                "sDom": "Tflt<'row DTTTFooter'<'col-sm-6'i><'col-sm-6'p>>",
                "oTableTools": {
                    "aButtons": [
                        //"copy",
                        //"print",                           
                        //{
                        //    "sExtends": "collection",
                        //    "sButtonText": "Save <i class=\"fa fa-angle-down\"></i>",
                        //    "aButtons": ["csv", "xls", "pdf"]
                        //}
                    ],
                    "sSwfPath": "assets/swf/copy_csv_xls_pdf.swf"
                },
                "language": {
                    "search": "",
                    "sLengthMenu": "_MENU_",
                    "oPaginate": {
                        "sPrevious": "Prev",
                        "sNext": "Next"
                    }
                },
                "aoColumns": [
                    { "bVisible": false },
                    null,
                    null,
                    null,
                    { "bSortable": false }
                ]
            });
        }
    };
}();

var InitiateSectionDataTable = function () {
    return {
        init: function () {
            //Datatable Initiating
            var oTable = $('#editabledatatable').dataTable({
                "aLengthMenu": [
                    [5, 10, 20, 100, -1],
                    [5, 10, 20, 100, "All"]
                ],
                "iDisplayLength": 10,
                "sPaginationType": "bootstrap",
                "sDom": "Tflt<'row DTTTFooter'<'col-sm-6'i><'col-sm-6'p>>",
                "oTableTools": {
                    "aButtons": [
                        ////"copy",
                        ////"print",                           
                        //{
                        //    "sExtends": "collection",
                        //    "sButtonText": "Save <i class=\"fa fa-angle-down\"></i>",
                        //    "aButtons": ["csv", "xls", "pdf"]
                        //}
                    ],
                    "sSwfPath": "assets/swf/copy_csv_xls_pdf.swf"
                },
                "language": {
                    "search": "",
                    "sLengthMenu": "_MENU_",
                    "oPaginate": {
                        "sPrevious": "Prev",
                        "sNext": "Next"
                    }
                },
                "aoColumns": [
                    { "bVisible": false },
                    null,
                    null,
                    null,
                    null,
                    { "bSortable": false }
                ]
            });
        }
    };
}();

var InitiateProjectsDataTable = function () {
    return {
        init: function () {
            //Datatable Initiating
            var oTable = $('#editabledatatable').dataTable({
                "aLengthMenu": [
                    [5, 10, 20, 100, -1],
                    [5, 10, 20, 100, "All"]
                ],
                "iDisplayLength": 10,
                "sPaginationType": "bootstrap",
                "sDom": "Tflt<'row DTTTFooter'<'col-sm-6'i><'col-sm-6'p>>",
                "oTableTools": {
                    "aButtons": [
                        ////"copy",
                        ////"print",                           
                        //{
                        //    "sExtends": "collection",
                        //    "sButtonText": "Save <i class=\"fa fa-angle-down\"></i>",
                        //    "aButtons": ["csv", "xls", "pdf"]
                        //}
                    ],
                    "sSwfPath": "assets/swf/copy_csv_xls_pdf.swf"
                },
                "language": {
                    "search": "",
                    "sLengthMenu": "_MENU_",
                    "oPaginate": {
                        "sPrevious": "Prev",
                        "sNext": "Next"
                    }
                },
                "aoColumns": [
                    { "bVisible": false },
                    null,
                    null,
                    null,
                    null,
                    null,
                    { "bVisible": false }
                ]
            });


            //Edit A Row
            $('#editabledatatable').on("click", 'a.edit', function (e) {
                e.preventDefault();
                var nRow = $(this).parents('tr')[0];
                var aData = oTable.fnGetData(nRow);
                $("#divnew").show();
                $(".topfocus").focus().css('border', '');//To be visible (scroll to top) the edit section by default
                $("#btnUpdate").show();
                $("#btnCreate").hide();
                $("#ancbacktoList").show();
                $("#ancaddNew").hide();
                $("#divGrid").hide();

                $("#divID").text(aData[0]);
                $("#ddlProjectType").val(aData[6]).prop('disabled', true);
                $("#ProjectCode").val(aData[2]).prop('disabled', true);
                $("#Project").val(aData[3]);
                if (aData[4] == "True")
                    $("#chk1").prop("checked", true);
                else
                    $("#chk1").prop("checked", false);
            });
        }
    };
}();

var InitiateCategoryDataTable = function () {
    return {
        init: function () {
            //Datatable Initiating
            var oTable = $('#editabledatatable').dataTable({
                "aLengthMenu": [
                    [5, 10, 20, 100, -1],
                    [5, 10, 20, 100, "All"]
                ],
                "iDisplayLength": 10,
                "sPaginationType": "bootstrap",
                "sDom": "Tflt<'row DTTTFooter'<'col-sm-6'i><'col-sm-6'p>>",
                "oTableTools": {
                    "aButtons": [
                        //"copy",
                        //"print",                           
                        //{
                        //    "sExtends": "collection",
                        //    "sButtonText": "Save <i class=\"fa fa-angle-down\"></i>",
                        //    "aButtons": ["csv", "xls", "pdf"]
                        //}
                    ],
                    "sSwfPath": "assets/swf/copy_csv_xls_pdf.swf"
                },
                "language": {
                    "search": "",
                    "sLengthMenu": "_MENU_",
                    "oPaginate": {
                        "sPrevious": "Prev",
                        "sNext": "Next"
                    }
                },
                "aoColumns": [
                    { "bVisible": false },
                    null,
                    null,
                    null,
                    null,
                    null
                ]
            });

            //Edit A Row
            $('#editabledatatable').on("click", 'a.edit', function (e) {
                e.preventDefault();
                var nRow = $(this).parents('tr')[0];
                var aData = oTable.fnGetData(nRow);
                $("#addDiv").show();
                $(".topfocus").focus().css('border', '');//To be visible (scroll to top) the edit section by default
                $("#btnUpdate").show();
                $("#btnCreate").hide();
                $("#ancbacktoList").show();
                $("#ancaddNew").hide();
                $("#divGrid").hide();

                $("#divID").text(aData[0]);
                $("#DocumentCategoryCode").val(aData[1]).prop('disabled', true);
                $("#DocumentCategoryName").val(aData[2]);
                $("#DocumentLevel").val(aData[3]).prop('disabled', true);
                //$("#FolderName").val(aData[4]);
                if (aData[4] == "True")
                    $("#chk1").prop("checked", true);
                else
                    $("#chk1").prop("checked", false);
            });
        }
    };
}();

var InitiateApprovalMatrixTable = function () {
    return {
        init: function () {
            //Datatable Initiating
            var oTable = $('#editabledatatable').dataTable({
                "aLengthMenu": [
                    [5, 10, 20, 100, -1],
                    [5, 10, 20, 100, "All"]
                ],
                "iDisplayLength": 10,
                "sPaginationType": "bootstrap",
                "sDom": "Tflt<'row DTTTFooter'<'col-sm-6'i><'col-sm-6'p>>",
                "oTableTools": {
                    "aButtons": [
                        //"copy",
                        //"print",                           
                        //{
                        //    "sExtends": "collection",
                        //    "sButtonText": "Save <i class=\"fa fa-angle-down\"></i>",
                        //    "aButtons": ["csv", "xls", "pdf"]
                        //}
                    ],
                    "sSwfPath": "assets/swf/copy_csv_xls_pdf.swf"
                },
                "language": {
                    "search": "",
                    "sLengthMenu": "_MENU_",
                    "oPaginate": {
                        "sPrevious": "Prev",
                        "sNext": "Next"
                    }
                },
                "aoColumns": [
                    { "bVisible": false },
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    { "bSortable": false }
                ]
            });

            var isEditing = null;

            //Delete an Existing Row
            $('#editabledatatable').on("click", 'a.delete', function (e) {
                e.preventDefault();
                var nRow = $(this).parents('tr')[0];
                var aData = oTable.fnGetData(nRow);
                var id = aData[0];
                //alert(id);
                //var projectTypeid = aData[8];
                //var projectid = aData[9];
                bootbox.confirm("Are you sure to Delete?", function (result) {
                    if (result) {
                        //oTable.fnDeleteRow(nRow);
                        setTimeout(function () {
                            DeleteApprovalMatrixStage(id, "/ApprovalMatrix/DeleteWFApprovalMatrix", "Stage", oTable, nRow);
                        }, 100);
                    }
                });
                return false;
            });

            //Edit A Row
            //$('#editabledatatable').on("click", 'a.edit', function (e) {
            //    e.preventDefault();
            //    var nRow = $(this).parents('tr')[0];
            //    var aData = oTable.fnGetData(nRow);
            //    var entID = aData[0];
            //    $("#divnew").show();
            //    $(".topfocus").focus().css('border', '');//To be visible (scroll to top) the edit section by default
            //    $("#btnUpdate").show();
            //    $("#btnCreate").hide();
            //    //Get the data for the ID and bind it to controls
            //    if (entID != "") {
            //        $.ajax({
            //            url: "/AppUser/GetDataForMasterID",
            //            dataType: "json",
            //            type: "POST",
            //            beforeSend: ShowProgressBar(),
            //            complete: HideProgressBar(),
            //            contentType: 'application/json; charset=utf-8',
            //            cache: false,
            //            data: JSON.stringify({ id: entID }),
            //            success: function (data) {
            //                retmsg = data.message;
            //                if (retmsg == '') {
            //                    $('.btn-danger').trigger('click');
            //                    $('.modal-body').html('Error while getting the data.');
            //                }
            //                else {
            //                    //alert(retmsg[1]);
            //                    $("#divID").html(retmsg.ID);
            //                    $("#txt1").val(retmsg.FirstName);
            //                    $("#txt2").val(retmsg.LastName);
            //                    $("#txt3").val(retmsg.Email);
            //                    $("#txt4").val(retmsg.Mobile);

            //                    //var dat = new Date(1501439400000);
            //                    //alert(dat);
            //                    //$('#id-date-picker-1').datepicker('setValue', dat);
            //                    //alert(ToJavaScriptDate(retmsg.DOB));
            //                    if (retmsg.DOB != null && retmsg.DOB != "")
            //                        $('#date1').datepicker('setValue', ToJavaScriptDate(retmsg.DOB));

            //                    $("#ddlMultiple").multipleSelect("setSelects", data.success);

            //                    if (retmsg.Gender == "Male")
            //                        $("#rbn1").prop("checked", true);
            //                    else
            //                        $("#rbn2").prop("checked", true);

            //                    if (retmsg.IsActive == true)
            //                        $("#chk1").prop('checked', true);//.attr("checked", true);
            //                    else
            //                        $("#chk1").prop('checked', false);//.removeAttr("checked");
            //                }
            //            },
            //            error: function (xhr) {
            //                $('.btn-danger').trigger('click');
            //                $('.modal-body').html('Error while getting the data.');
            //            }
            //        });
            //    }

            //});
        }
    };
}();


var InitiateProjectUsersTable = function () {
    return {
        init: function () {
            //Datatable Initiating
            var oTable = $('#editabledatatable').dataTable({
                "aLengthMenu": [
                    [5, 10, 20, 100, -1],
                    [5, 10, 20, 100, "All"]
                ],
                "iDisplayLength": 10,
                "sPaginationType": "bootstrap",
                "sDom": "Tflt<'row DTTTFooter'<'col-sm-6'i><'col-sm-6'p>>",
                "oTableTools": {
                    "aButtons": [
                        //"copy",
                        //"print",                           
                        //{
                        //    "sExtends": "collection",
                        //    "sButtonText": "Save <i class=\"fa fa-angle-down\"></i>",
                        //    "aButtons": ["csv", "xls", "pdf"]
                        //}
                    ],
                    "sSwfPath": "assets/swf/copy_csv_xls_pdf.swf"
                },
                "language": {
                    "search": "",
                    "sLengthMenu": "_MENU_",
                    "oPaginate": {
                        "sPrevious": "Prev",
                        "sNext": "Next"
                    }
                },
                "aoColumns": [
                    { "bVisible": false },
                    null,
                    null,
                    null,
                    null,
                    null,
                    { "bVisible": false },
                    { "bVisible": false },
                    { "bVisible": false }
                ]
            });

            var isEditing = null;

            //Delete an Existing Row
            $('#editabledatatable').on("click", 'a.delete', function (e) {
                e.preventDefault();
                var nRow = $(this).parents('tr')[0];
                var aData = oTable.fnGetData(nRow);
                var id = aData[0];
                var projectTypeid = aData[7];
                var projectid = aData[8];
                bootbox.confirm("Are you sure to Delete?", function (result) {
                    if (result) {
                        //oTable.fnDeleteRow(nRow);
                        setTimeout(function () {
                            DeleteProjectUser(id, projectTypeid, projectid, "/ProjectUsers/DeleteProjectUser", "User", oTable, nRow);
                        }, 100);                        
                    }
                });
                return false;
            });

            //Edit A Row
            $('#editabledatatable').on("click", 'a.edit', function (e) {
                e.preventDefault();
                var nRow = $(this).parents('tr')[0];
                var aData = oTable.fnGetData(nRow);
                var entID = aData[0];
                $("#divnew").show();
                $(".topfocus").focus().css('border', '');//To be visible (scroll to top) the edit section by default
                $("#btnUpdate").show();
                $("#btnCreate").hide();
                //Get the data for the ID and bind it to controls
                if (entID != "") {
                    $.ajax({
                        url: "/AppUser/GetDataForMasterID",
                        dataType: "json",
                        type: "POST",
                        beforeSend: ShowProgressBar(),
                        complete: HideProgressBar(),
                        contentType: 'application/json; charset=utf-8',
                        cache: false,
                        data: JSON.stringify({ id: entID }),
                        success: function (data) {
                            retmsg = data.message;
                            if (retmsg == '') {
                                $('.btn-danger').trigger('click');
                                $('.modal-body').html('Error while getting the data.');
                            }
                            else {
                                //alert(retmsg[1]);
                                $("#divID").html(retmsg.ID);
                                $("#txt1").val(retmsg.FirstName);
                                $("#txt2").val(retmsg.LastName);
                                $("#txt3").val(retmsg.Email);
                                $("#txt4").val(retmsg.Mobile);

                                //var dat = new Date(1501439400000);
                                //alert(dat);
                                //$('#id-date-picker-1').datepicker('setValue', dat);
                                //alert(ToJavaScriptDate(retmsg.DOB));
                                if (retmsg.DOB != null && retmsg.DOB != "")
                                    $('#date1').datepicker('setValue', ToJavaScriptDate(retmsg.DOB));

                                $("#ddlMultiple").multipleSelect("setSelects", data.success);

                                if (retmsg.Gender == "Male")
                                    $("#rbn1").prop("checked", true);
                                else
                                    $("#rbn2").prop("checked", true);

                                if (retmsg.IsActive == true)
                                    $("#chk1").prop('checked', true);//.attr("checked", true);
                                else
                                    $("#chk1").prop('checked', false);//.removeAttr("checked");
                            }
                        },
                        error: function (xhr) {
                            $('.btn-danger').trigger('click');
                            $('.modal-body').html('Error while getting the data.');
                        }
                    });
                }

            });
        }
    };
}();

function DeleteApprovalMatrixStage(id, cururl, entity, oTable, nRow) {
    $.ajax({
        url: cururl,
        dataType: "json",
        type: "POST",
        beforeSend: ShowProgressBar(),
        complete: HideProgressBar(),
        contentType: 'application/json; charset=utf-8',
        cache: false,
        data: JSON.stringify({ ID: id}),
        success: function (data) {
            if (data.success) {
                var incStr = data.message.includes("sucessfully")
                //alert(incStr);
                if (incStr) {
                    oTable.fnDeleteRow(nRow);
                }
                $('.btn-success').trigger('click');
                $('.modal-body').html(data.message);
            }
            else {
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

function DeleteProjectUser(id, projectTypeid, projectid, cururl, entity, oTable, nRow) {
    $.ajax({
        url: cururl,
        dataType: "json",
        type: "POST",
        beforeSend: ShowProgressBar(),
        complete: HideProgressBar(),
        contentType: 'application/json; charset=utf-8',
        cache: false,
        data: JSON.stringify({ UserID: id, ProjectTypeID: projectTypeid, ProjectID: projectid }),
        success: function (data) {
            if (data.success) {
                var incStr = data.message.includes("sucessfully")
                //alert(incStr);
                if (incStr)
                {
                    oTable.fnDeleteRow(nRow);
                }

                $('.btn-success').trigger('click');
                $('.modal-body').html(data.message);
                
            }
            else {
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

function DeleteEntityItem(id, projectTypeid, projectid, cururl, entity) {
    $.ajax({
        url: cururl,
        dataType: "json",
        type: "POST",
        beforeSend: ShowProgressBar(),
        complete: HideProgressBar(),
        contentType: 'application/json; charset=utf-8',
        cache: false,
        data: JSON.stringify({ UserID: id, ProjectTypeID: projectTypeid, ProjectID: projectid }),
        success: function (data) {
            if (data.success) {
                $('.btn-success').trigger('click');
                $('.modal-body').html(data.message);
                retmsg = data.message;
            }
            else {
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

var InitiateUserApprovalMatrixTable_1 = function () {
    return {
        init: function () {
            //Datatable Initiating
            var oTable = $('#editabledatatable').dataTable({
                "aLengthMenu": [
                    [5, 10, 20, 100, -1],
                    [5, 10, 20, 100, "All"]
                ],
                "iDisplayLength": 100,
                "sPaginationType": "bootstrap",
                "sDom": "Tflt<'row DTTTFooter'<'col-sm-6'i><'col-sm-6'p>>",
                "oTableTools": {
                    "aButtons": [
                        //"copy",
                        //"print",                           
                        //{
                        //    "sExtends": "collection",
                        //    "sButtonText": "Save <i class=\"fa fa-angle-down\"></i>",
                        //    "aButtons": ["csv", "xls", "pdf"]
                        //}
                    ],
                    "sSwfPath": "assets/swf/copy_csv_xls_pdf.swf"
                },
                "language": {
                    "search": "",
                    "sLengthMenu": "_MENU_",
                    "oPaginate": {
                        "sPrevious": "Prev",
                        "sNext": "Next"
                    }
                },
                "aoColumns": [
                    { "bVisible": false },
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    { "bSortable": false }
                ]
            });            
        }
    };
}();

var InitiateUserApprovalMatrixTable = function () {
    return {
        init: function () {
            //Datatable Initiating
            var oTable = $('#editabledatatable').dataTable({
                "aLengthMenu": [
                    [10, 20, 50, 100, -1],
                    [10, 20, 50, 100, "All"]
                ],
                "iDisplayLength": 50,
                "sPaginationType": "bootstrap",
                "sDom": "Tflt<'row DTTTFooter'<'col-sm-6'i><'col-sm-6'p>>",
                "oTableTools": {
                    "aButtons": [
                        //"copy",
                        //"print",                           
                        //{
                        //    "sExtends": "collection",
                        //    "sButtonText": "Save <i class=\"fa fa-angle-down\"></i>",
                        //    "aButtons": ["csv", "xls", "pdf"]
                        //}
                    ],
                    "sSwfPath": "assets/swf/copy_csv_xls_pdf.swf"
                },
                "language": {
                    "search": "",
                    "sLengthMenu": "_MENU_",
                    "oPaginate": {
                        "sPrevious": "Prev",
                        "sNext": "Next"
                    }
                },
                "aoColumns": [
                    { "bVisible": false },
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    { "bSortable": false },
                    { "bVisible": false },
                    { "bVisible": false },
                    { "bVisible": false },
                    { "bVisible": false }
                ]
            });

        }
    };
}();


var InitiateWFItemsDataTable = function () {
    return {
        init: function () {
            //Datatable Initiating
            var oTable = $('#editabledatatableWFItems').dataTable({
                "aLengthMenu": [
                    [5, 10, 20, 100, -1],
                    [5, 10, 20, 100, "All"]
                ],
                "iDisplayLength": 10,
                "sPaginationType": "bootstrap",
                "sDom": "Tflt<'row DTTTFooter'<'col-sm-6'i><'col-sm-6'p>>",
                "oTableTools": {
                    "aButtons": [
                        //"copy",
                        //"print",                           
                        //{
                        //    "sExtends": "collection",
                        //    "sButtonText": "Save <i class=\"fa fa-angle-down\"></i>",
                        //    "aButtons": ["csv", "xls", "pdf"]
                        //}
                    ],
                    "sSwfPath": "assets/swf/copy_csv_xls_pdf.swf"
                },
                "language": {
                    "search": "",
                    "sLengthMenu": "_MENU_",
                    "oPaginate": {
                        "sPrevious": "Prev",
                        "sNext": "Next"
                    }
                },
                "aoColumns": [
                    { "bVisible": false },
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null
                ]
            });            
        }
    };
}();

var InitiateWFItemsDataTable_1 = function () {
    return {
        init: function () {
            //Datatable Initiating
            var oTable = $('#editabledatatableWFItems_1').dataTable({
                "aLengthMenu": [
                    [5, 10, 20, 100, -1],
                    [5, 10, 20, 100, "All"]
                ],
                "iDisplayLength": 10,
                "sPaginationType": "bootstrap",
                "sDom": "Tflt<'row DTTTFooter'<'col-sm-6'i><'col-sm-6'p>>",
                "oTableTools": {
                    "aButtons": [
                        //"copy",
                        //"print",                           
                        //{
                        //    "sExtends": "collection",
                        //    "sButtonText": "Save <i class=\"fa fa-angle-down\"></i>",
                        //    "aButtons": ["csv", "xls", "pdf"]
                        //}
                    ],
                    "sSwfPath": "assets/swf/copy_csv_xls_pdf.swf"
                },
                "language": {
                    "search": "",
                    "sLengthMenu": "_MENU_",
                    "oPaginate": {
                        "sPrevious": "Prev",
                        "sNext": "Next"
                    }
                },
                "aoColumns": [
                    { "bVisible": false },
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null
                ]
            });
        }
    };
}();


var InitiateSystemUsersTable = function () {
    return {
        init: function () {
            //Datatable Initiating
            var oTable = $('#editabledatatable').dataTable({
                "aLengthMenu": [
                    [5, 10, 20, 100, -1],
                    [5, 10, 20, 100, "All"]
                ],
                "iDisplayLength": 10,
                "sPaginationType": "bootstrap",
                "sDom": "Tflt<'row DTTTFooter'<'col-sm-6'i><'col-sm-6'p>>",
                "oTableTools": {
                    "aButtons": [
                        //"copy",
                        //"print",                           
                        //{
                        //    "sExtends": "collection",
                        //    "sButtonText": "Save <i class=\"fa fa-angle-down\"></i>",
                        //    "aButtons": ["csv", "xls", "pdf"]
                        //}
                    ],
                    "sSwfPath": "assets/swf/copy_csv_xls_pdf.swf"
                },
                "language": {
                    "search": "",
                    "sLengthMenu": "_MENU_",
                    "oPaginate": {
                        "sPrevious": "Prev",
                        "sNext": "Next"
                    }
                },
                "aoColumns": [
                    { "bVisible": false },
                    null,
                    null,
                    null,
                    null,
                    null,
                    { "bSortable": false }
                ]
            });

            var isEditing = null;

            //Delete an Existing Row
            $('#editabledatatable').on("click", 'a.delete', function (e) {
                e.preventDefault();
                var nRow = $(this).parents('tr')[0];
                var aData = oTable.fnGetData(nRow);
                var id = aData[0];
                bootbox.confirm("Are you sure to Delete?", function (result) {
                    if (result) {
                        //oTable.fnDeleteRow(nRow);
                        setTimeout(function () {
                            DeleteUser(id, "/Users/DeleteUser", "User", oTable, nRow);
                        }, 100);
                    }
                });
                return false;
            });

            //Edit A Row
            $('#editabledatatable').on("click", 'a.edit', function (e) {
                e.preventDefault();
                var nRow = $(this).parents('tr')[0];
                var aData = oTable.fnGetData(nRow);
                var entID = aData[0];                
                $("#divnew").show();
                $(".topfocus").focus().css('border', '');//To be visible (scroll to top) the edit section by default
                $("#btnUpdate").show();
                $("#btnCreate").hide();
                $("#divIsActive").show();
                $("#divGrid").hide();
                //Get the data for the ID and bind it to controls
                if (entID != "") {
                    $.ajax({
                        url: "/Users/GetUserData",
                        dataType: "json",
                        type: "POST",
                        beforeSend: ShowProgressBar(),
                        complete: HideProgressBar(),
                        contentType: 'application/json; charset=utf-8',
                        cache: false,
                        data: JSON.stringify({ id: entID }),
                        success: function (data) {
                            retmsg = data.message;
                            if (retmsg == '') {
                                $('.btn-danger').trigger('click');
                                $('.modal-body').html('Error while getting the data.');
                            }
                            else {
                                $("#divID").html(retmsg[0].ID);
                                $("#LoginID").val(retmsg[0].LoginID);//.prop("disabled","disabled");
                                $("#DisplayName").val(retmsg[0].DisplayName).prop("disabled", "disabled");
                                $("#EmailID").val(retmsg[0].EmailID);
                                if (retmsg[0].IsQMSAdmin == true)
                                    $("#chk2").prop('checked', true);
                                else
                                    $("#chk2").prop('checked', false);
                                if (retmsg[0].IsActive == true)
                                    $("#chk1").prop('checked', true);
                                else
                                    $("#chk1").prop('checked', false);
                            }
                        },
                        error: function (xhr) {
                            $('.btn-danger').trigger('click');
                            $('.modal-body').html('Error while getting the data.');
                        }
                    });
                }

            });
        }
    };
}();


function DeleteUser(id, cururl, entity, oTable, nRow) {
    $.ajax({
        url: cururl,
        dataType: "json",
        type: "POST",
        beforeSend: ShowProgressBar(),
        complete: HideProgressBar(),
        contentType: 'application/json; charset=utf-8',
        cache: false,
        data: JSON.stringify({ UserID: id }),
        success: function (data) {
            if (data.success) {
                var incStr = data.message.includes("sucessfully")
                //alert(incStr);
                if (incStr) {
                    oTable.fnDeleteRow(nRow);
                }
                $('.btn-success').trigger('click');
                $('.modal-body').html(data.message);
            }
            else {
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


var InitiateDocHistoryDataTable = function () {
    return {
        init: function () {
            //Datatable Initiating
            var oTable = $('#editabledatatable').dataTable({
                "aLengthMenu": [
                    [5, 10, 20, 100, -1],
                    [5, 10, 20, 100, "All"]
                ],
                "iDisplayLength": 10,
                "sPaginationType": "bootstrap",
                "sDom": "Tflt<'row DTTTFooter'<'col-sm-6'i><'col-sm-6'p>>",
                "oTableTools": {
                    "aButtons": [
                        //"copy",
                        //"print",                           
                        //{
                        //    "sExtends": "collection",
                        //    "sButtonText": "Save <i class=\"fa fa-angle-down\"></i>",
                        //    "aButtons": ["csv", "xls", "pdf"]
                        //}
                    ],
                    "sSwfPath": "assets/swf/copy_csv_xls_pdf.swf"
                },
                "language": {
                    "search": "",
                    "sLengthMenu": "_MENU_",
                    "oPaginate": {
                        "sPrevious": "Prev",
                        "sNext": "Next"
                    }
                },
                "aoColumns": [
                    { "bVisible": false },
                    { "bVisible": false },
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null
                ]
            });
        }
    };
}();

var InitiateExternalDocumentDataTable = function () {
    return {
        init: function () {
            //Datatable Initiating
            var oTable = $('#editabledatatable').dataTable({
                "aLengthMenu": [
                    [5, 10, 20, 100, -1],
                    [5, 10, 20, 100, "All"]
                ],
                "iDisplayLength": 10,
                "sPaginationType": "bootstrap",
                "sDom": "Tflt<'row DTTTFooter'<'col-sm-6'i><'col-sm-6'p>>",
                "oTableTools": {
                    "aButtons": [
                        //"copy",
                        //"print",                           
                        //{
                        //    "sExtends": "collection",
                        //    "sButtonText": "Save <i class=\"fa fa-angle-down\"></i>",
                        //    "aButtons": ["csv", "xls", "pdf"]
                        //}
                    ],
                    "sSwfPath": "assets/swf/copy_csv_xls_pdf.swf"
                },
                "language": {
                    "search": "",
                    "sLengthMenu": "_MENU_",
                    "oPaginate": {
                        "sPrevious": "Prev",
                        "sNext": "Next"
                    }
                },
                "aoColumns": [
                    { "bVisible": false },
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    { "bSortable": false }

                ]
            });
        }
    };
}();

var InitiateExternalDocumentsDataTable = function () {
    return {
        init: function () {
            //Datatable Initiating
            var oTable = $('#editabledatatable').dataTable({
                "aLengthMenu": [
                    [5, 10, 20, 100, -1],
                    [5, 10, 20, 100, "All"]
                ],
                "iDisplayLength": 10,
                "sPaginationType": "bootstrap",
                "sDom": "Tflt<'row DTTTFooter'<'col-sm-6'i><'col-sm-6'p>>",
                "oTableTools": {
                    "aButtons": [
                        //"copy",
                        //"print",                           
                        //{
                        //    "sExtends": "collection",
                        //    "sButtonText": "Save <i class=\"fa fa-angle-down\"></i>",
                        //    "aButtons": ["csv", "xls", "pdf"]
                        //}
                    ],
                    "sSwfPath": "assets/swf/copy_csv_xls_pdf.swf"
                },
                "language": {
                    "search": "",
                    "sLengthMenu": "_MENU_",
                    "oPaginate": {
                        "sPrevious": "Prev",
                        "sNext": "Next"
                    }
                },
                "aoColumns": [
                    { "bVisible": false },
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null,
                    null

                ]
            });
        }
    };
}();


var InitiateExtCategoryDataTable = function () {
    return {
        init: function () {
            //Datatable Initiating
            var oTable = $('#editabledatatable').dataTable({
                "aLengthMenu": [
                    [5, 10, 20, 100, -1],
                    [5, 10, 20, 100, "All"]
                ],
                "iDisplayLength": 10,
                "sPaginationType": "bootstrap",
                "sDom": "Tflt<'row DTTTFooter'<'col-sm-6'i><'col-sm-6'p>>",
                "oTableTools": {
                    "aButtons": [
                        //"copy",
                        //"print",                           
                        //{
                        //    "sExtends": "collection",
                        //    "sButtonText": "Save <i class=\"fa fa-angle-down\"></i>",
                        //    "aButtons": ["csv", "xls", "pdf"]
                        //}
                    ],
                    "sSwfPath": "assets/swf/copy_csv_xls_pdf.swf"
                },
                "language": {
                    "search": "",
                    "sLengthMenu": "_MENU_",
                    "oPaginate": {
                        "sPrevious": "Prev",
                        "sNext": "Next"
                    }
                },
                "aoColumns": [
                    { "bVisible": false },
                    null,
                    null,
                    { "bSortable": false }
                ]
            });
        }
    };
}();

var InitiateExtSubCategoryDataTable = function () {
    return {
        init: function () {
            //Datatable Initiating
            var oTable = $('#editabledatatable').dataTable({
                "aLengthMenu": [
                    [5, 10, 20, 100, -1],
                    [5, 10, 20, 100, "All"]
                ],
                "iDisplayLength": 10,
                "sPaginationType": "bootstrap",
                "sDom": "Tflt<'row DTTTFooter'<'col-sm-6'i><'col-sm-6'p>>",
                "oTableTools": {
                    "aButtons": [
                        //"copy",
                        //"print",                           
                        //{
                        //    "sExtends": "collection",
                        //    "sButtonText": "Save <i class=\"fa fa-angle-down\"></i>",
                        //    "aButtons": ["csv", "xls", "pdf"]
                        //}
                    ],
                    "sSwfPath": "assets/swf/copy_csv_xls_pdf.swf"
                },
                "language": {
                    "search": "",
                    "sLengthMenu": "_MENU_",
                    "oPaginate": {
                        "sPrevious": "Prev",
                        "sNext": "Next"
                    }
                },
                "aoColumns": [
                    { "bVisible": false },
                    null,
                    null,
                    null,
                    { "bSortable": false }
                ]
            });
        }
    };
}();


