console.log(".");

function getCookie(name) {
    const cookies = document.cookie.split(';');
    for (let cookie of cookies) {
        let [key, value] = cookie.trim().split('=');
        if (key === name) {
            return decodeURIComponent(value);
        }
    }
    return null;
}

function updateICBObjectInfo(objectNo, storeNo, data) {
    const url = "/OtherApp/ICB/Ashx/Handler.ashx";
    const params = $.extend({
        "action": "update_object_info",
        "object_no": objectNo,
        "store_no": storeNo
    }, data || {});

    return $.ajax({
        url: url,
        data: params,
        type: "post",
        cache: false,
        dataType: "json"
    });
}

function updateICBObjectInfoAndRedirect(objectNo, storeNo, agreeCB, redirectUrl) {
    updateICBObjectInfo(objectNo, storeNo, {
        "agree_cb": agreeCB
    }).always(function () {
        window.location.href = redirectUrl;
    });
}

//聯賣
function ICBEvent() {

    var fInfo = {
        ICBStoreAgents: [],
        StoreNo: ""
    };

    fInfo.StoreNo = getCookie('store_id');
    //function setICBStoreAgents(storeNo) {

    //    const url = "/OtherApp/ICB/Ashx/Handler.ashx";
    //    const params = {
    //        "action": "get_store_agents"
    //        , "store_no": storeNo
    //    };
    //    $.ajax({
    //        url: url,
    //        data: params,
    //        type: "get",
    //        async: true,
    //        cache: false,
    //        dataType: "json",
    //        success: function (res) {
    //            if (res == null || res.length == 0) {
    //                fInfo.ICBStoreAgents = [];
    //                setICBAgentOptions();
    //                return;
    //            }
    //            fInfo.ICBStoreAgents = res;
    //            setICBAgentOptions();


    //        },
    //        error: function () {
    //            return;
    //        }
    //    });
    //}
    function setICBObjectStatus(objectNo, storeNo) {
        const url = "/OtherApp/ICB/Ashx/Handler.ashx";
        const params = {
            "action": "get_object_info",
            "object_no": objectNo,
            "store_no": storeNo
        };

        $.ajax({
            url: url,
            data: params,
            type: "get",
            async: true,
            cache: false,
            dataType: "json",
            success: function (res) {
                const $statusSpan = $(`.col_icb_agree_cb_status[object_no='${objectNo}'][store_no='${storeNo}']`);
                //const $empSpan = $(`.col_icb_emp_name[object_no='${objectNo}'][store_no='${storeNo}']`);

                if (!res) {
                    $statusSpan.text("不同意");
                    //$empSpan.text("未指定");
                    return;
                }

                // 設定顯示內容
                $statusSpan.text(res.agree_cb ? "同意" : "不同意");
                //$empSpan.text(res.emp_name || "未指定");
            },
            error: function () {
                console.error("取得物件資訊失敗：" + objectNo);
            }
        });
    }

    //function setICBAgentOptions() {

    //    let $item = $("#icb_emp_no");
    //    if (fInfo.ICBStoreAgents == null) {
    //        $item.html("");
    //        return;
    //    }

    //    let code = "<option value=''>-請選擇人員-</option>";

    //    const items = fInfo.ICBStoreAgents;
    //    const count = items.length;

    //    for (let i = 0; i < count; i++) {
    //        code += `<option value='${items[i].agent_no}'>${items[i].agent_name}</option>`;
    //    }

    //    $item.html(code);
    //}


    $("#select_all_icb_item").on("click", function () {
        var checkboxes = document.querySelectorAll('input[name="check_cb"]'); // 取得所有資料列的 CheckBox

        for (var i = 0; i < checkboxes.length; i++) {
            checkboxes[i].checked = this.checked; // 將所有資料列的 CheckBox 狀態設定為全選 CheckBox 的狀態
        }
    });

    $("#button_icb_set_agree_cb").on("click", function (e) {
        e.preventDefault();
        var status = $("#icb_agree_cb option:selected").text();
        var agreeCB = $("#icb_agree_cb").val();
        if (agreeCB == "") {
            new Dialog().Alert("請先選擇聯賣狀態!!");
            return;
        }
        const len = $('input[name="check_cb"]:checked').size();
        if (len == 0) {
            new Dialog().Alert("請先勾選聯賣物件!!");
            return;
        }
        var status = $("#icb_agree_cb option:selected").text();
        new Dialog().ConfirmV2(`已選取項目的聯賣狀態，確定要設定為「${status}」!!`, () => {

            var promiseList = [];

            $('input[name="check_cb"]:checked').each(function () {
                var objectNo = $(this).attr('object_no');
                var storeNo = $(this).attr('store_no');

                // 包裹一層 Promise，並帶入 objectNo 與 storeNo
                const p = updateICBObjectInfo(objectNo, storeNo, {
                    "agree_cb": agreeCB
                }).then(
                    res => ({ status: "fulfilled", objectNo, storeNo, result: res }),
                    err => ({ status: "rejected", objectNo, storeNo, error: err })
                );

                promiseList.push(p);
            });

            Promise.all(promiseList).then(results => {
                const failed = results.filter(x => x.status === "rejected");

                if (failed.length > 0) {
                    let msg = `有 ${failed.length} 筆資料更新失敗：\n\n`;
                    failed.forEach(f => {
                        msg += `物件編號：${f.objectNo}, 店代號：${f.storeNo}\n`;
                    });
                    new Dialog().Alert(msg);
                } else {
                    new Dialog().Alert("資料更新成功！");
                }


                // 無論成功或失敗都更新顯示
                $(".col_icb_agree_cb_status").each(function () {
                    const objectNo = $(this).attr("object_no");
                    const storeNo = $(this).attr("store_no");
                    setICBObjectStatus(objectNo, storeNo);
                });
            });
        });


    });

    $("#button_icb_set_emp_no").on("click", function (e) {

        e.preventDefault();
        var empName = $("#icb_emp_no option:selected").text();
        var empNo = $("#icb_emp_no").val();
        if (empNo == "") {
            new Dialog().Alert("請先選擇負責人員!!");
            return;
        }
        const len = $('input[name="check_cb"]:checked').size();
        if (len == 0) {
            new Dialog().Alert("請先勾選聯賣物件!!");
            return;
        }
        new Dialog().ConfirmV2(`已選取項目的聯賣物件負責人，確定要設定為「${empName}」!!`, () => {

            var promiseList = [];

            $('input[name="check_cb"]:checked').each(function () {
                var objectNo = $(this).attr('object_no');
                var storeNo = $(this).attr('store_no');

                // 包裹一層 Promise，並帶入 objectNo 與 storeNo
                const p = updateICBObjectInfo(objectNo, storeNo, {
                    "emp_no": empNo
                }).then(
                    res => ({ status: "fulfilled", objectNo, storeNo, result: res }),
                    err => ({ status: "rejected", objectNo, storeNo, error: err })
                );

                promiseList.push(p);
            });

            Promise.all(promiseList).then(results => {
                const failed = results.filter(x => x.status === "rejected");

                if (failed.length > 0) {
                    let msg = `有 ${failed.length} 筆資料更新失敗：\n\n`;
                    failed.forEach(f => {
                        msg += `物件編號：${f.objectNo}, 店代號：${f.storeNo}\n`;
                    });
                    new Dialog().Alert(msg);
                } else {
                    new Dialog().Alert("資料更新成功！");
                }


                // 無論成功或失敗都更新顯示
                $(".col_icb_agree_cb_status").each(function () {
                    const objectNo = $(this).attr("object_no");
                    const storeNo = $(this).attr("store_no");
                    setICBObjectStatus(objectNo, storeNo);
                });
            });
        });
    });

    //setICBStoreAgents(fInfo.StoreNo);

    $(".col_icb_agree_cb_status").each(function () {
        const objectNo = $(this).attr("object_no");
        const storeNo = $(this).attr("store_no");
        setICBObjectStatus(objectNo, storeNo);
    });

    $(document).on("click", ".div_check_cb", function (e) {
        // 避免點到 checkbox 本身觸發兩次
        if ($(e.target).is("input[type='checkbox']")) {
            return;
        }
        e.preventDefault();
        const checkbox = $(this).find("input[type='checkbox']");
        checkbox.prop("checked", !checkbox.prop("checked"));
    });
    
}


function onSub() {
    //alert("a")
    var i;
    currentAjax_a = $.ajax({
        url: "../B_CustomerManage/Get_customer.aspx", //url,
        dataType: "json",
        type: "post",
        //data:"submit=true",
        beforeSend: function (bs) {
            $('#oksend_r_store2').html("<div style='margin:0 auto;width:100px'><img src='../images/LodingPic.gif' /></div>");
        },
        success: function (res) {
            console.log(res); //return;
            var _total = res.length;

            $('#oksend_r_store2').html('');
            $("#search_count").text(_total);
            $('#oksend_result').show();

            if (_total == 0) { return; }
            store_json = res;
            toPage(1);
        }
    });
}

// hide & show message- blockUI-查詢
$(document).ready(function () {
    $('#ImageButton1').click(function () {
        $("#loading").show();
        /*
        $.blockUI({
        message: $('#welcome'),
        background: '#FFF'
        });
        setTimeout($.unblockUI, 8000);*/
    });

    ICBEvent();
});

//報表
//    $(document).ready(function () {
//        $('#GridView1_objectlist').click(function () {
//            $("#loading").show();
//            /*
//            $.blockUI({
//            message: $('#welcome'),
//            background: '#FFF'
//            });
//            setTimeout($.unblockUI, 8000);*/
//        });
//    });

//輸入檢查,數字及小數點
function ValidateNumber(e, pnumber) {
    if (!/^\d+[.]?\d*$/.test(pnumber)) {
        var newValue = /^\d+[.]?\d*/.exec(e.value);
        if (newValue != null) {
            e.value = newValue;
        }
        else {
            e.value = "";
        }
    }
    return false;
}
function toPage(p) {
    var _html = _url = '',
        _html2 = '1',
        n = 0,
        _total = store_json.length,
        _limit = 13,
        _pages = Math.ceil(_total / _limit);
    b = (p - 1) * _limit;
    c = b + (_limit - 1),
        prev_page = (p > 1) ? (p - 1) : 1,
        next_page = ((p + 1) <= _pages) ? (p + 1) : _pages;

    if (c > _total) {
        c = _total - 1;
    } else if (c < _limit && _total < _limit) {
        c = _total - 1;
    }

    _html = '<table width="100%" border="0" cellpadding="1" cellspacing="1" bgcolor="#f5f5f5" class="inpage_content_font_02"><tr class="inpage_content_font_01"><td height="30" align="center" bgcolor="#f96f00"><center>案名</center></td><td height="30" align="center" bgcolor="#f96f00">物件編號</td><td height="30" align="center" bgcolor="#f96f00">物件編號1</td></tr>';
    for (var i = b; i <= c; i++) {
        //_url = 'https://superweb3.etwarm.com.tw/new_town/index.php?storeid=' + store_json[i].id;
        _url = 'https://store.etwarm.com.tw/' + store_json[i].id;
        _html += '<tr>';
        //_html += '<td height="30" align="center" bgcolor="#FFFFFF">' + store_json[i].id + '<br></td>';
        _html += '<td height="30" align="center" bgcolor="#FFFFFF">' + store_json[i].id + '<br></td>';
        _html += '<td height="30" align="center" bgcolor="#FFFFFF">' + store_json[i].name + '<br></td>';
        _html += '<td height="30" align="center" bgcolor="#FFFFFF">' + store_json[i].sex + '<br></td>';
        _html += '</tr>';
        n++;
    }
    _html += '</table></div>';
    $('#oksend_r_store2').html(_html);

    if (_pages > 1) {
        _html2 = "<a href='javascript:void(0);' onclick='toPage(1)'>第一頁</a> ";
        _html2 += "<a href='javascript:void(0);' onclick='toPage(" + prev_page + ")'>上一頁</a> ";
        _html2 += p + "/" + _pages;
        _html2 += " <a href='javascript:void(0);' onclick='toPage(" + next_page + ")'>下一頁</a> ";
        _html2 += "<a href='javascript:void(0);' onclick='toPage(" + _pages + ")'>最後頁</a>";
    }

    $('#store_page').html(_html2);
}
$(function () {

    $("#carspacedialog").dialog({
        autoOpen: false,
        width: 600
    });
    $("#dialopener").click(function () {


        $("#carspacedialog").dialog("open");
    });
});

$('#chkAll').on('change', function () {
    var checked = $(this).prop('checked');
    if (checked == true) {
        $('#hidChk').val('true');
    } else {
        $('#hidChk').val('false');
    }

    $(".GridViewItemStyle input.check-item").each(function () {
        var _checked = $(this).prop('checked');
        if (checked != _checked) {
            //$(this).click();
            $(this).prop('checked', checked).change();
        }
    });
});
$(".GridViewItemStyle input:not([name='check_cb'])").on('change', function () {
    var $this = $(this);
    var checked = $this.prop('checked');
    var id = $this.attr('id');
    var HidID = id.replace('ckbID', 'HidID'); //物件編號
    var HidIDs = id.replace('ckbID', 'HidIDs'); //店代號
    var HidSrc = id.replace('ckbID', 'HidSrc'); //來源

    var HidID_val = $('#' + HidID).val();
    var HidIDs_val = $('#' + HidIDs).val();
    var HidSrc_val = $('#' + HidSrc).val();

    var hidAll = $('#hidAll').val();
    var hidAlls = $('#hidAlls').val();
    var hidAllSrc = $('#hidAllSrc').val();

    var hidAll_arr = [];
    var hidAlls_arr = [];
    var hidAllSrc_arr = [];

    if (hidAll != '') { hidAll_arr = hidAll.split(','); }
    if (hidAlls != '') { hidAlls_arr = hidAlls.split(','); }
    if (hidAllSrc != '') { hidAllSrc_arr = hidAllSrc.split(','); }
    if (checked == true) {
        hidAll_arr.push(HidID_val);
        hidAlls_arr.push(HidIDs_val);
        hidAllSrc_arr.push(HidSrc_val);
        $this.closest('tr').css("background", '#d8fffe');
    } else {
        var len = hidAll_arr.length || 0;
        var _index = $.inArray(HidID_val, hidAll_arr);
        if (_index >= 0) {
            hidAll_arr.splice(_index, 1);
            hidAlls_arr.splice(_index, 1);
            hidAllSrc_arr.splice(_index, 1);
        }
        $this.closest('tr').css("background", '#ffffff');
    }
    $('#hidAll').val(hidAll_arr.join(','));
    $('#hidAlls').val(hidAlls_arr.join(','));
    $('#hidAllSrc').val(hidAllSrc_arr.join(','));
    $('#hidAllst').val($(".GridViewItemStyle input").filter(":checked").eq(0).closest('td').find('span').eq(0).text());
});

function GB_showCenter(v, url, height, width) {
    $.colorbox({
        'width': width,
        'height': height,
        'iframe': true,
        'href': url,
        'fixed': true
    });
}
$('#GridView1_ImgBtnFieldRpt2').click(function () {
    var url = "../I_SystemSetting/report_set.aspx?source=top&reporttype=A";
    $.colorbox({
        'width': '100%',
        'innerHeight': $(window).height() * 8 / 10,
        'scrolling': true,
        'iframe': true,
        'href': url,
        'fixed': true
    })

});
