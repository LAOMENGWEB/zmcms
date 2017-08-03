function DoLocation(a) {
    a != lastCtrl && (lastCtrl.className = ""),
    a.className = "hover",
    lastCtrl = a
}
function setLeft(a) {
    parent.leftFrame.location = a
}

//检查登陆
function checklogin(a) {
    if ("" == $.trim(a.t0.value)) return $.message({
        content: "\u8bf7\u8f93\u5165\u7528\u6237\u540d"
    }),
    a.t0.focus(),
    !1;
    if ("" == $.trim(a.t1.value)) return $.message({
        content: "\u8bf7\u8f93\u5165\u5bc6\u7801"
    }),
    a.t1.focus(),
    !1;
    if ("" == $.trim(a.t2.value)) return $.message({
        content: "\u8bf7\u8f93\u5165\u9a8c\u8bc1\u7801"
    }),
    a.t2.focus(),
    !1;
    var b, c;
    return b = "login.asp?act=check",
    c = "t0=" + encodeURIComponent($.trim(a.t0.value)),
    c += "&t1=" + encodeURIComponent($.trim(a.t1.value)),
    c += "&t2=" + encodeURIComponent($.trim(a.t2.value)),
    $.ajax({
        type: "post",
        cache: !1,
        url: b,
        data: c,
        error: function() {
            $.message({
                type: "error",
                content: "\u670d\u52a1\u5668\u9519\u8bef\uff0c\u64cd\u4f5c\u5931\u8d25\uff01"
            })
        },
        success: function(a) {
            var b = a.substring(0, 1),
            c = a.substring(1);
            switch (b) {
            case "0":
                $.message({
                    type:
                    "error",
                    content: c
                });
                break;
            case "1":
                $.message({
                    type:
                    "ok",
                    content: c
                }),
                setTimeout("location.href='./'", 1500);
                break;
            default:
                $.message({
                    type:
                    "error",
                    content: a,
                    time: 1e4
                })
            }
        }
    }),
    !1
}
$(function() {
    $(".addlist dt").mouseover(function() {
        $(this).css("background", "#EDF3FA")
    }),
    $(".addlist dt").mouseout(function() {
        $(this).css("background", "#F5F8FC")
    }),
    $(".addlist dt input").focus(function() {
        $(this).addClass("input")
    }),
    $(".addlist dt input").blur(function() {
        $(this).removeClass("input")
    }),
    $(".addlist dt textarea").focus(function() {
        $(this).addClass("input")
    }),
    $(".addlist dt textarea").blur(function() {
        $(this).removeClass("input")
    }),
    $("#table tr").mouseover(function() {
        $(this).css("background", "#EBF4FC")
    }),
    $("#table tr").mouseout(function() {
        $(this).css("background", "#FFF")
    }),
    $("#table input").focus(function() {
        $(this).addClass("input")
    }),
    $("#table input").blur(function() {
        $(this).removeClass("input")
    }),
    $(".dropdown").hover(function() {
        $("ul", this).show()
    },
    function() {
        $("ul", this).hide()
    }),
    $("#color").click(function() {
        var a = $(this).width();
        $(this).height();
        var c = $(this).offset().top,
        d = $(this).offset().left;
        $.fn.colorpicker({
            x: c,
            y: d + a,
            inputid: "s2",
            backid: "color"
        })
    }),
    $(".bnt").click(function() {
        var a = $(this).attr("config"),
        b = a.split(":"),
        c = b[0],
        d = b[1],
        e = b[2],
        f = b[3],
        g = $.dialog.through;
        g({
            title: "图片上传",
            content: "<iframe src='/plugins/WeiXin/lib/upload.asp?filetype=" + d + "&cid=" + e + "&classid=" + f + "' width='560' height='400' frameborder='0' id='swfupload'></iframe>",
            padding: "10px",
            lock: !0,
            opacity: "0.5",
            ok: function() {
                var a = $("#att-status", art.dialog.top.$("#swfupload")[0].contentWindow.document).html();
                if (null != a) {
                    var b, d;
                    b = a.substring(0, 1),
                    d = a.substring(1),
                    $("input[name=" + c + "]").attr("value", d)
                }
            },
            cancelVal: "取消",
            cancel: !0
        })
    })
});
var lastCtrl = new Object;