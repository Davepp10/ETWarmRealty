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

$(document).ready(function () {
    $("#ImageButton1").click(function () {
        $("#loading").show();
    });
});
