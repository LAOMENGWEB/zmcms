/// <summary>ShowWin</summary>
/// 汉字转拼音
/// </summary>
/// <param name="UrlPath"></param>
/// <param name="objId">要转换文本框中的文字ID</param>
/// <param name="objVal">返回到文本框的ID</param>
/// <param name="objIdType">拼音类型：1=转拼音首字母,2=转拼音全拼音</param>
/// <param name="startStr">追加开始字符串</param>
/// <param name="endStr">追加结束字符串</param>
function GetPy(UrlPath, objId, objVal, objIdType, startStr, endStr) {
    if ($('#' + objId).val() != '') {
        var TipType = parseInt(objIdType);
        var TipVal = $('#' + objVal + '_' + TipType).val();
        $('#' + objVal + '_' + TipType).val('转换中……');
        if (TipType == 6) {
            GetPySzm(objId, objVal, startStr, endStr);
        } else {
            GetPyQpy(objId, objVal, startStr, endStr);
        }
        $('#' + objVal + '_' + TipType).val(TipVal);
    }
}

// 汉字转拼音首字母
function GetPySzm(objId, objVal, startStr, endStr) {
    if ($('#' + objId).val() != '') {
        $('#' + objVal).val(startStr + Pinyin.GetJP($('#' + objId).val()) + endStr);
    }
}
// 汉字转拼音全拼音
function GetPyQpy(objId, objVal, startStr, endStr) {
    if ($('#' + objId).val() != '') {
        $('#' + objVal).val(startStr + Pinyin.GetQP($('#' + objId).val()) + endStr);
    }
}