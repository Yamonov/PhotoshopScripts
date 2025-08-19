/*
<javascriptresource>
<name>InDesignに合わせてリサイズ</name>
<category>YPresets</category>
</javascriptresource>
*/

#target photoshop

// 設定値（ターゲットPPIリスト、警告用しきい値）
var targetPPIList = [350, 400, 600, 1200];
var defaultTargetPPI = 350;
var efScaleMin = 0.9;
var efScaleMax = 1.1;
var scaleMax = 2.5;

function main() {
    // Photoshopでアクティブなドキュメントの情報を取得
    if (!app.documents.length) {
        alert("開いているドキュメントがありません。");
        return;
    }
    var doc = app.activeDocument;
    var imgPath = doc.fullName.fsName;
    var docWidthPx = doc.width.as("px");
    var docHeightPx = doc.height.as("px");
    var currentPPI = doc.resolution;

    // InDesignにBridgeTalk経由でリンク画像情報を問い合わせ
    var bt = new BridgeTalk();
    bt.target = "indesign";
    bt.body = "(" + inDesignSide.toString() + ")(" + encodeURI(imgPath).toSource() + ");";

    bt.onResult = function (res) {
        var data = res.body;
        if (!data || data == "null") {
            alert("InDesignで該当リンク画像が見つかりません。");
            return;
        }
        var obj = eval('(' + data + ')');
        var absScale = obj.absScale; // 最大 absoluteHorizontalScale (%)
        var linkStatus = obj.linkStatus; // 0:正常、1:リンク切れ、2:未更新

        if (linkStatus === 1) {
            alert("リンク切れ画像です。");
            return;
        } else if (linkStatus === 2) {
            alert("リンクが更新されていません。");
            return;
        }

        // InDesign配置画像の長辺サイズ（mm）を計算
        var longPx = Math.max(docWidthPx, docHeightPx);
        var placedLongMM = longPx * (absScale / 100) / currentPPI * 25.4;

        // 指定したPPIで必要な長辺ピクセル数を計算
        function calcRequiredPx(longMM, ppi) {
            return Math.round(longMM * ppi / 25.4);
        }

        // リサイズ確認用ダイアログを表示
        var messageBase = "InDesign配置サイズ(長辺): " + placedLongMM.toFixed(2) + " mm\n"
            + "現在ppi: " + currentPPI + "\n"
            + "配置スケール: " + absScale.toFixed(3) + " %\n"
            + "画像ピクセル: " + longPx + "\n";

        var dialogResult = showConfirmDialog(messageBase, placedLongMM, longPx, imgPath);
        if (dialogResult === null) return;
        var targetPPI = dialogResult.ppi;
        var upscaleMethod = dialogResult.method;
        var downscaleMethod = dialogResult.downMethod;

        // 必要な長辺ピクセル数を計算
        var newLongPx = calcRequiredPx(placedLongMM, targetPPI);
        var scaleRatio = newLongPx / longPx;

        var newWidthPx, newHeightPx;
        if (docWidthPx >= docHeightPx) {
            newWidthPx = newLongPx;
            newHeightPx = Math.round(docHeightPx * scaleRatio);
        } else {
            newHeightPx = newLongPx;
            newWidthPx = Math.round(docWidthPx * scaleRatio);
        }

        // 画像のリサイズ処理
        if (scaleRatio < 1) {
            // 縮小
            if (downscaleMethod === "bicubic") {
                doc.resizeImage(UnitValue(newWidthPx, "px"), UnitValue(newHeightPx, "px"), targetPPI, ResampleMethod.BICUBIC);
            } else if (downscaleMethod === "nearestNeighbor") {
                doc.resizeImage(UnitValue(newWidthPx, "px"), UnitValue(newHeightPx, "px"), targetPPI, ResampleMethod.NEARESTNEIGHBOR);
            } else {
                doc.resizeImage(UnitValue(newWidthPx, "px"), UnitValue(newHeightPx, "px"), targetPPI, ResampleMethod.BICUBIC);
            }
        } else {
            // 拡大
            if (upscaleMethod === "deepUpscale") {
                var desc = new ActionDescriptor();
                desc.putUnitDouble(charIDToTypeID('Wdth'), charIDToTypeID('#Pxl'), newWidthPx);
                desc.putUnitDouble(charIDToTypeID('Hght'), charIDToTypeID('#Pxl'), newHeightPx);
                desc.putUnitDouble(charIDToTypeID('Rslt'), charIDToTypeID('#Rsl'), targetPPI);
                desc.putBoolean(stringIDToTypeID('scaleStyles'), true);
                desc.putEnumerated(charIDToTypeID('Intr'), charIDToTypeID('Intp'), stringIDToTypeID('deepUpscale'));
                executeAction(charIDToTypeID('ImgS'), desc, DialogModes.NO);
            } else if (upscaleMethod === "preserveDetails") {
                doc.resizeImage(UnitValue(newWidthPx, "px"), UnitValue(newHeightPx, "px"), targetPPI, ResampleMethod.PRESERVEDETAILS);
            } else if (upscaleMethod === "nearestNeighbor") {
                doc.resizeImage(UnitValue(newWidthPx, "px"), UnitValue(newHeightPx, "px"), targetPPI, ResampleMethod.NEARESTNEIGHBOR);
            } else {
                doc.resizeImage(UnitValue(newWidthPx, "px"), UnitValue(newHeightPx, "px"), targetPPI, ResampleMethod.BICUBIC);
            }
        }
    };

    bt.onError = function (e) {
        alert("InDesign通信エラー: " + e.body);
    };
    bt.send(30);
}

// InDesignのリンク情報から、該当ファイルの最大スケール値とリンク状態を返す
function inDesignSide(imgPath) {
    var decodedPath = decodeURI(imgPath);
    if (app.documents.length === 0) return null;
    var doc = app.activeDocument;
    var links = doc.links;
    var mlMax = 0;
    var found = false;
    var linkStat = -1; // 0:正常, 1:切れ, 2:未更新, -1:該当なし
    for (var i = 0; i < links.length; i++) {
        var link = links[i];
        if (link.filePath == decodedPath) {
            found = true;
            var parent = link.parent;
            if (link.status == LinkStatus.NORMAL) {
                var ml = parent.absoluteHorizontalScale;
                if (ml > mlMax) {
                    mlMax = ml;
                    linkStat = 0;
                }
            } else if (link.status == LinkStatus.LINK_MISSING) {
                linkStat = 1;
            } else {
                linkStat = 2;
            }
        }
    }
    if (!found) return null; // 一致するリンクなし
    // 正常以外はmlMaxが0でも返す
    return '{"absScale":'+mlMax+',"linkStatus":'+linkStat+'}';
}

// リサイズ確認ダイアログのUI定義
function showConfirmDialog(messageBase, placedLongMM, longPx, imgPath) {
    var dlg = new Window("dialog", "画像リサイズ確認");
    dlg.orientation = "column";
    dlg.alignChildren = ["fill", "top"];
    dlg.margins = 15;

    var infoPanel = dlg.add("panel", undefined, "画像情報");
    infoPanel.orientation = "column";
    infoPanel.alignChildren = ["fill", "top"];
    infoPanel.margins = [10,12,10,12];
    var msgText = infoPanel.add("statictext", undefined, messageBase, {multiline:true});
    msgText.minimumSize.width = 400;
    try { msgText.graphics.font = ScriptUI.newFont(msgText.graphics.font.name, 'Bold', msgText.graphics.font.size); } catch(e){}

    var radioGrp = dlg.add("group");
    radioGrp.orientation = "row";
    radioGrp.alignChildren = ["left", "center"];
    radioGrp.margins = [0, 10, 0, 10];
    radioGrp.add("statictext", undefined, "指定解像度:");
    var radioButtons = [];
    var selectedIndex = 0;
    for (var i = 0; i < targetPPIList.length; i++) {
        var rb = radioGrp.add("radiobutton", undefined, String(targetPPIList[i]));
        radioButtons.push(rb);
        if (targetPPIList[i] === defaultTargetPPI) selectedIndex = i;
    }
    radioButtons[selectedIndex].value = true;

    var methodPanel = dlg.add("panel", undefined, "拡大メソッド");
    methodPanel.orientation = "row";
    methodPanel.alignChildren = ["left", "center"];
    methodPanel.margins = [10, 12, 10, 12];
    var methodRadioButtons = [];
    var methodLabels = [
        "ディテールを保持2.0（推奨）",
        "ディテールを保持（旧）",
        "ニアレストネイバー"
    ];
    var methodValues = [
        "deepUpscale",
        "preserveDetails",
        "nearestNeighbor"
    ];
    var defaultMethodIndex = 0;
    for (var i = 0; i < methodLabels.length; i++) {
        var rb = methodPanel.add("radiobutton", undefined, methodLabels[i]);
        methodRadioButtons.push(rb);
        if (i === defaultMethodIndex) rb.value = true;
    }

    var downMethodPanel = dlg.add("panel", undefined, "縮小メソッド");
    downMethodPanel.orientation = "row";
    downMethodPanel.alignChildren = ["left", "center"];
    downMethodPanel.margins = [10, 12, 10, 12];
    var downMethodRadioButtons = [];
    var downMethodLabels = [
        "バイキュービック（滑らか）",
        "ニアレストネイバー"
    ];
    var downMethodValues = [
        "bicubic",
        "nearestNeighbor"
    ];
    var defaultDownMethodIndex = 0;
    for (var i = 0; i < downMethodLabels.length; i++) {
        var rb = downMethodPanel.add("radiobutton", undefined, downMethodLabels[i]);
        downMethodRadioButtons.push(rb);
        if (i === defaultDownMethodIndex) rb.value = true;
    }

    var scaleText = dlg.add("statictext", undefined, "");
    scaleText.minimumSize.width = 400;
    try { scaleText.graphics.font = ScriptUI.newFont(scaleText.graphics.font.name, 'Bold', scaleText.graphics.font.size); } catch(e){}

    var warnText = dlg.add("statictext", undefined, "", {multiline:true});
    warnText.minimumSize.width = 400;
    warnText.maximumSize.width = 400;

    function updateUI() {
        var selectedPPI = null;
        for (var i = 0; i < radioButtons.length; i++) {
            if (radioButtons[i].value) { selectedPPI = parseInt(radioButtons[i].text); break; }
        }
        var requiredPx = Math.round(placedLongMM * selectedPPI / 25.4);
        var scale = requiredPx / longPx;
        var scalePct = scale * 100;
        scaleText.text = "拡縮率: " + scalePct.toFixed(2) + " %";
        var warnMsg = "";
        var warnColor = null;
        if (scale >= efScaleMin && scale <= efScaleMax) {
            warnMsg = "※警告: 拡大率が " + scalePct.toFixed(2) + "% です。無駄な拡縮の可能性があります。";
            warnColor = [1, 0.4, 0, 1];
        }
        if (scale > scaleMax) {
            warnMsg += (warnMsg ? "\n" : "") + "！: 拡大率が " + (scaleMax * 100).toFixed(0) + "% を超えています。Photoshop以外の手段を検討してください。";
            warnColor = [1, 0, 0, 1];
        }
        warnText.text = warnMsg;
        if (warnColor) warnText.graphics.foregroundColor = warnText.graphics.newPen(warnText.graphics.PenType.SOLID_COLOR, warnColor, 1);
        else warnText.graphics.foregroundColor = warnText.graphics.newPen(warnText.graphics.PenType.SOLID_COLOR, [0,0,0,1], 1);
    }
    updateUI();
    for (var i = 0; i < radioButtons.length; i++) { radioButtons[i].onClick = updateUI; }

    var btnGrp = dlg.add("group");
    btnGrp.alignment = "center";
    btnGrp.margins = [0,15,0,0];
    var okBtn = btnGrp.add("button", undefined, "処理する");
    var cancelBtn = btnGrp.add("button", undefined, "キャンセル");

    okBtn.onClick = function() { dlg.close(1); };
    cancelBtn.onClick = function() { dlg.close(0); };

    var result = dlg.show() === 1 ? (function() {
        var ppi = null, method = null, downMethod = null;
        for (var i = 0; i < radioButtons.length; i++) if (radioButtons[i].value) ppi = parseInt(radioButtons[i].text);
        for (var i = 0; i < methodRadioButtons.length; i++) if (methodRadioButtons[i].value) method = methodValues[i];
        for (var i = 0; i < downMethodRadioButtons.length; i++) if (downMethodRadioButtons[i].value) downMethod = downMethodValues[i];
        return {ppi: ppi, method: method, downMethod: downMethod};
    })() : null;
    return result;
}

// メイン処理の実行
main();