//auto size header width based on data rows width
function resizecolumns(headerid, tableid, outerdivid) {
    try {
        var headerRow = $("[id$=" + headerid + "] tr").eq(0);
        if ($("[id$=" + tableid + "] tr").length == 0) {
            var width = $("#" + outerdivid).innerWidth() / (headerRow.children("th").length);
            $.each(headerRow.children("th"), function (idx, obj) {
                obj.width = width;
            });
        } else {
            $.each($("[id$=" + tableid + "] tr").eq(0).children("td"), function (idx, obj) {
                var width = obj.getBoundingClientRect().width;
                obj.width = width;
                headerRow.children("th")[idx].width = width + 2;
            });
        }
    } catch (e) {
        //do nothing
    }

}


function LoadMoreRecords(divid, gridviewid, serviceURL, webMethodURL, pIndex, pSize, pCount, loaderdivid, languageid, message) {
    $("#" + divid).on("scroll", function (e) {
        var $o = $(e.currentTarget);
        if ($o[0].scrollHeight - $o.scrollTop() - 15 <= $o.outerHeight()) {
            pIndex++;
            GetRecords(gridviewid, loaderdivid, serviceURL + "offset=" + ((pIndex - 1) * pSize).toString() + "&pagesize=" + pSize.toString(), webMethodURL, pIndex, pCount, languageid, message);
        }
    });
}

//Function to make AJAX call to the Web Method
function GetRecords(gridviewid, loaderdivid, serviceURL, webMethodURL, pIndex, pCount, languageid, message) {
    if (pIndex <= pCount) {
        $.ajax({
            type: "POST",
            url: webMethodURL,
            async: true,
            data: "{ URL: \"" + serviceURL + "\",LanguageID:" + languageid + " }",
            contentType: "application/json; charset=utf-8",
            dataType: "json"
        })
        .done(function (response) {
            OnSuccess(response, gridviewid, loaderdivid);
        })
        .fail(function (response) {
            $(loaderdivid).html('<center>' + response.d + '</center>');
        });
    }
    else {
        $(loaderdivid).html('<center>' + message + '</center>');
    }
}