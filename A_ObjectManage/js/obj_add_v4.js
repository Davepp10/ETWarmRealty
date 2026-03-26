function getUrlParameter(name) {
    var search = window.location.search.substring(1);
    if (!search) {
        return "";
    }

    var params = search.split("&");
    for (var i = 0; i < params.length; i++) {
        var pair = params[i].split("=");
        if (decodeURIComponent(pair[0]) === name) {
            return pair.length > 1 ? decodeURIComponent(pair[1].replace(/\+/g, " ")) : "";
        }
    }

    return "";
}

function updateICBObjectInfo(objectNo, storeNo, data) {
    var url = "/OtherApp/ICB/Ashx/Handler.ashx";
    var params = $.extend({
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

$(document).ready(function () {
    $("#ImageButton1").click(function () {
        $("#loading").show();
    });
});
