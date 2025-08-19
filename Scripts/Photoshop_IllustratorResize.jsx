/*
<javascriptresource>
<name>Illustratorに合わせてリサイズ</name>
<category>YPresets</category>
</javascriptresource>
*/

#target photoshop

// Photoshopから実行するスクリプト

// 拡大率・縮小率の閾値をAppleScript仕様に合わせて変数で設定
var efScaleMin = 0.9; // 90% 以下で警告表示の下限
var efScaleMax = 1.1; // 110% 以上で警告表示の上限
var scaleMax = 2.5;   // 250% を超える拡大は警告

// 解像度リスト・初期値（ユーザー選択用）
var targetPPIList = [350, 400, 600, 1200];
var defaultTargetPPI = 350;

// メイン処理関数
function main() {
    // ドキュメントが開かれていなければ処理を停止
    if (!app.documents.length) {
        alert("開いているドキュメントがありません。");
        return;
    }

    var activeDoc = app.activeDocument;

    // 現在のドキュメントのファイルパスを取得（ファイルシステムのパス形式）
    var imgPath = activeDoc.fullName.fsName;

    // BridgeTalkでIllustrator側に画像パスを渡し、配置画像情報を取得する
    var bt = new BridgeTalk();
    bt.target = "illustrator";

    // Illustrator側関数を文字列化し、imgPathをURIエンコードして渡す
    bt.body =
        "(" + illustratorSideForPT.toString() + ")(" + encodeURI(imgPath).toSource() + ");";

    // Illustratorからの応答受信時の処理
    bt.onResult = function (res) {
        var data = res.body;
        if (data == "null" || !data) {
            alert("Illustratorドキュメント内に一致するリンク画像が見つかりません。");
            return;
        }
        // JSON文字列をオブジェクトに変換
        var obj = JSON.parse(data);
        var longSidePt = obj.longSidePt; // 配置画像の長辺サイズ（pt単位）
        var linkStatus = obj.linkStatus; // リンク状態（0=正常,1=リンク切れ,2=未更新）

        // リンク状態によるエラーチェック
        if (linkStatus === 1) {
            alert("リンク切れの画像が見つかりました。処理を中止します。");
            return;
        } else if (linkStatus === 2) {
            alert("未更新のリンク画像があります。処理を中止します。");
            return;
        }

        // Illustratorのポイント(pt)単位をミリメートル(mm)に変換（1インチ=25.4mm、1インチ=72pt）
        var longSideMM = longSidePt * 25.4 / 72;

        // Photoshopドキュメントの長辺ピクセル数を取得
        var docWidthPx = activeDoc.width.as("px");
        var docHeightPx = activeDoc.height.as("px");
        var longSidePx = Math.max(docWidthPx, docHeightPx);

        // Photoshopの現在の解像度（ppi）を取得
        var currentPPI = activeDoc.resolution;

        // Illustrator上の配置画像のppiを計算（ppi = px / inch = px / (mm / 25.4)）
        var placedPPI = longSidePx / (longSideMM / 25.4);

        // ScriptUIで解像度選択と拡縮率確認のダイアログを表示し、ユーザーの選択を取得
        function showConfirmDialog(messageBase, placedPPI, imgPath) {
            // ダイアログの生成と設定
            var dlg = new Window("dialog", "画像リサイズ確認");
            dlg.orientation = "column";
            dlg.alignChildren = ["fill", "top"];
            dlg.margins = 15;

            // 画像情報表示用パネル作成
            var infoPanel = dlg.add("panel", undefined, "画像情報");
            infoPanel.orientation = "column";
            infoPanel.alignChildren = ["fill", "top"];
            infoPanel.margins = [10,12,10,12];

            // メッセージテキスト表示（配置サイズとppiなど）
            var msgText = infoPanel.add("statictext", undefined, messageBase, {multiline:true});
            msgText.minimumSize.width = 400;
            try {
                var fnt = msgText.graphics.font;
                msgText.graphics.font = ScriptUI.newFont(fnt.name, 'Bold', fnt.size);
            } catch(e) {}

            // 解像度選択用ラジオボタン群を横並びで作成
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
                if (targetPPIList[i] === defaultTargetPPI) {
                    selectedIndex = i;
                }
            }
            radioButtons[selectedIndex].value = true;

            // --- 拡大メソッド選択ラジオボタン群追加 ---
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
                if (i === defaultMethodIndex) {
                    rb.value = true;
                }
            }

            // --- 縮小メソッド選択ラジオボタン群追加 ---
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
                if (i === defaultDownMethodIndex) {
                    rb.value = true;
                }
            }

            // 拡縮率表示用テキスト
            var scaleText = dlg.add("statictext", undefined, "");
            scaleText.minimumSize.width = 400;
            try {
                var fnt = scaleText.graphics.font;
                scaleText.graphics.font = ScriptUI.newFont(fnt.name, 'Bold', fnt.size);
            } catch(e) {}

            // 警告表示用テキスト（複数行可・常に全体表示）
            var warnText = dlg.add("edittext", undefined, "", {multiline:true, readonly:true});
            warnText.preferredSize = [400, 80];
            warnText.minimumSize.width = 400;
            warnText.maximumSize.width = 400;
            warnText.alignment = ["fill", "top"]; // 横幅は親に合わせ、上寄せ

            // ダイアログの推奨サイズを指定（全コントロールが収まるように）
            dlg.preferredSize = [460, 360];

            // UI更新関数：選択解像度に基づき拡縮率と警告文を更新
            function updateUI() {
                var selectedPPI = null;
                for (var i = 0; i < radioButtons.length; i++) {
                    if (radioButtons[i].value) {
                        selectedPPI = parseInt(radioButtons[i].text);
                        break;
                    }
                }
                var scale = (selectedPPI / placedPPI);
                var scalePct = scale * 100;
                scaleText.text = "拡縮率: " + scalePct.toFixed(2) + " %";

                // 拡縮率に応じた警告メッセージ
                var warnMsg = "";
                if (scale >= efScaleMin && scale <= efScaleMax) {
                    warnMsg = "※警告: 拡大率が " + scalePct.toFixed(2) + "% です。無駄な拡縮の可能性があります。";
                }
                if (scale > scaleMax) {
                    warnMsg += (warnMsg ? "\n" : "") + "！: 拡大率が " + (scaleMax * 100).toFixed(0) + "% を超えています。Photoshop以外の手段を検討してください。";
                }
                warnText.text = warnMsg;
                dlg.layout.layout(true);
                dlg.layout.resize();
            }

            // 初期化とラジオボタンのクリックイベントにUI更新関数を割り当て
            updateUI();
            for (var i = 0; i < radioButtons.length; i++) {
                radioButtons[i].onClick = updateUI;
            }

            // ボタン配置グループ作成（中央揃え）
            var btnGrp = dlg.add("group");
            btnGrp.orientation = "row";
            btnGrp.alignment = "center";          // グループ自体を中央揃え
            btnGrp.alignChildren = ["center","center"]; // ボタンも中央揃え
            btnGrp.margins = [0,15,0,0];
            var okBtn = btnGrp.add("button", undefined, "処理する (S / Enter)");
            var cancelBtn = btnGrp.add("button", undefined, "キャンセル (Esc)");
            var aiButton = btnGrp.add("button", undefined, "Illustratorで選択 (I)");

            // ヘルプチップ（ツールチップ）
            okBtn.helpTip = "S キー または Enter で実行";
            cancelBtn.helpTip = "Esc でキャンセル";
            aiButton.helpTip = "I キー で Illustrator で選択";

            // 既定キー設定（EnterでOK、EscでCancel）
            okBtn.properties = {name: 'ok'};       // Enter = OK
            cancelBtn.properties = {name: 'cancel'}; // Esc = Cancel
            okBtn.active = true;                    // 初期フォーカスをOKに

            // 「処理する」ボタン：ダイアログを閉じてOKを返す
            okBtn.onClick = function() {
                dlg.close(1);
            };
            // 「キャンセル」ボタン：ダイアログを閉じてキャンセルを返す
            cancelBtn.onClick = function() {
                dlg.close(0);
            };
            // 「Illustratorで選択」ボタン：Illustrator側で該当配置画像を選択する処理をBridgeTalkで送信
            aiButton.onClick = function() {
                dlg.close();
                var filePath = imgPath;

                // Illustrator側で該当PlacedItemを再帰的に収集する関数
                function collectPlacedItems(pageItem, decodedPath, result) {
                    if (pageItem.typename === "PlacedItem" && pageItem.file && pageItem.file.fsName == decodedPath) {
                        result.push(pageItem);
                        return;
                    }
                    if (pageItem.pageItems && pageItem.pageItems.length > 0) {
                        for (var i = 0; i < pageItem.pageItems.length; i++) {
                            collectPlacedItems(pageItem.pageItems[i], decodedPath, result);
                        }
                    }
                }

                // Illustrator側で該当画像を選択し、ズームする関数（ズーム処理は削除済み）
                function illustratorSelectAndZoomFitByPath(imgPath) {
                    var decodedPath = decodeURI(imgPath);
                    if (app.documents.length === 0) return;
                    var doc = app.activeDocument;
                    doc.selection = null;
                    var foundItems = [];
                    // ドキュメント全体から再帰的に配置画像を探す
                    for (var i = 0; i < doc.pageItems.length; i++) {
                        collectPlacedItems(doc.pageItems[i], decodedPath, foundItems);
                    }
                    if (!foundItems.length) return;
                    // 見つかった配置画像を全て選択状態にする
                    for (var i = 0; i < foundItems.length; i++) {
                        foundItems[i].selected = true;
                    }
                    // ズーム処理は削除済み
                    doc.activate();
                    try { app.redraw(); } catch (e) {}
                    try { app.activate(); } catch (e) {}
                }

                // BridgeTalkでIllustrator側に選択処理を送信
                var bt = new BridgeTalk();
                bt.target = "illustrator";
                bt.body =
                    "(" +
                    "function() {" +
                        collectPlacedItems.toString() + "\n" +
                        illustratorSelectAndZoomFitByPath.toString() + "\n" +
                        "illustratorSelectAndZoomFitByPath('" + encodeURI(filePath).replace(/'/g,"\\'") + "');" +
                    "}" +
                    ")();";
                bt.onError = function(e) { alert("Illustrator BridgeTalk Error: " + e.body); };
                bt.send();

                // 少し待ってから（Photoshop側のモーダル解放待ち）、AppleScriptで前面化を強制
                try {
                    $.sleep(200); // 200ms 程度の待機で十分なことが多い
                    // macOSのosascriptで Illustrator を前面化
                    // ダブルクォートのエスケープに注意
                    system.callSystem('osascript -e "tell application \\"Adobe Illustrator 2025\\" to activate"');
                } catch (e) {
                    // 無視
                }
            };

            // キーボードショートカット（キャンセル以外）
            // S: 処理する / I: Illustratorで選択
            dlg.addEventListener('keydown', function(k) {
                try {
                    var n = String(k.keyName || '').toUpperCase();
                    if (n === 'S') { okBtn.notify('onClick'); k.preventDefault(); }
                    else if (n === 'I') { aiButton.notify('onClick'); k.preventDefault(); }
                } catch (e) {}
            });

            // 初期表示時にレイアウトを確定させ、ボタンが隠れないようにする
            dlg.onShow = function() {
                try {
                    dlg.layout.layout(true);
                    dlg.layout.resize();
                    dlg.center();
                } catch (e) {}
            };
            // ダイアログ表示し、OKなら選択された解像度と拡大メソッド・縮小メソッドを返す。キャンセルならnullを返す。
            var result = dlg.show() === 1 ? (function() {
                var ppi = null;
                for (var i = 0; i < radioButtons.length; i++) {
                    if (radioButtons[i].value) {
                        ppi = parseInt(radioButtons[i].text);
                        break;
                    }
                }
                var method = null;
                for (var i = 0; i < methodRadioButtons.length; i++) {
                    if (methodRadioButtons[i].value) {
                        method = methodValues[i];
                        break;
                    }
                }
                var downMethod = null;
                for (var i = 0; i < downMethodRadioButtons.length; i++) {
                    if (downMethodRadioButtons[i].value) {
                        downMethod = downMethodValues[i];
                        break;
                    }
                }
                return {ppi: ppi, method: method, downMethod: downMethod};
            })() : null;
            return result;
        }

        // ダイアログに表示する基本メッセージ（配置サイズとppi）
        var messageBase = "配置サイズ（長辺）: " + longSideMM.toFixed(2) + " mm\n" +
                          "配置画像のppi: " + placedPPI.toFixed(2) + "\n";

        // 解像度選択ダイアログを表示し、ユーザーの選択を取得
        var dialogResult = showConfirmDialog(messageBase, placedPPI, imgPath);

        // キャンセル時は処理を中止
        if (dialogResult === null) {
            return;
        }

        var targetPPI = dialogResult.ppi;
        var upscaleMethod = dialogResult.method;
        var downscaleMethod = dialogResult.downMethod;
        var scaleRatio = targetPPI / placedPPI;

        // 長辺を拡縮率に応じてリサイズ（縦横比維持）
        var newWidthPx, newHeightPx;
        if (docWidthPx >= docHeightPx) {
            newWidthPx = longSidePx * scaleRatio;
            newHeightPx = docHeightPx * scaleRatio;
        } else {
            newHeightPx = longSidePx * scaleRatio;
            newWidthPx = docWidthPx * scaleRatio;
        }

        // 拡縮率が1未満なら選択された縮小メソッドで縮小
        if (scaleRatio < 1) {
            if (downscaleMethod === "bicubic") {
                activeDoc.resizeImage(UnitValue(newWidthPx, "px"), UnitValue(newHeightPx, "px"), targetPPI, ResampleMethod.BICUBIC);
            } else if (downscaleMethod === "nearestNeighbor") {
                activeDoc.resizeImage(UnitValue(newWidthPx, "px"), UnitValue(newHeightPx, "px"), targetPPI, ResampleMethod.NEARESTNEIGHBOR);
            } else {
                // 万一未知の値の場合はBICUBIC
                activeDoc.resizeImage(UnitValue(newWidthPx, "px"), UnitValue(newHeightPx, "px"), targetPPI, ResampleMethod.BICUBIC);
            }
        } else {
            // 拡大時は選択された拡大メソッドで分岐
            if (upscaleMethod === "deepUpscale") {
                // ActionDescriptorでディテールを保持2.0（推奨）
                var desc = new ActionDescriptor();
                desc.putUnitDouble(charIDToTypeID('Wdth'), charIDToTypeID('#Pxl'), newWidthPx);
                desc.putUnitDouble(charIDToTypeID('Hght'), charIDToTypeID('#Pxl'), newHeightPx);
                desc.putUnitDouble(charIDToTypeID('Rslt'), charIDToTypeID('#Rsl'), targetPPI);
                desc.putBoolean(stringIDToTypeID('scaleStyles'), true);
                desc.putEnumerated(charIDToTypeID('Intr'), charIDToTypeID('Intp'), stringIDToTypeID('deepUpscale'));
                executeAction(charIDToTypeID('ImgS'), desc, DialogModes.NO);
            } else if (upscaleMethod === "preserveDetails") {
                // ResampleMethod.PRESERVEDETAILS
                activeDoc.resizeImage(UnitValue(newWidthPx, "px"), UnitValue(newHeightPx, "px"), targetPPI, ResampleMethod.PRESERVEDETAILS);
            } else if (upscaleMethod === "nearestNeighbor") {
                // ResampleMethod.NEARESTNEIGHBOR
                activeDoc.resizeImage(UnitValue(newWidthPx, "px"), UnitValue(newHeightPx, "px"), targetPPI, ResampleMethod.NEARESTNEIGHBOR);
            } else {
                // 万一未知の値の場合はBICUBIC
                activeDoc.resizeImage(UnitValue(newWidthPx, "px"), UnitValue(newHeightPx, "px"), targetPPI, ResampleMethod.BICUBIC);
            }
        }
    };

    // BridgeTalk通信エラー時の処理
    bt.onError = function (e) {
        alert("Illustratorとの通信エラー: " + e.body);
    };

    // BridgeTalk送信（30秒タイムアウト）
    bt.send(30);
}

// Illustrator側関数：指定画像パスの配置画像の長辺(pt)とリンク状態を返す
// 入力: Photoshopから渡された画像パス（URIエンコード済み）
// 出力: JSON文字列 { longSidePt: 数値, linkStatus: 0|1|2 }
function illustratorSideForPT(psImgPath) {
    var decodedPath = decodeURI(psImgPath);

    // ドキュメントが開かれていなければnull返却
    if (app.documents.length === 0) return null;

    // 再帰的に配置画像を収集する関数
    function collectPlacedItems(pageItem, decodedPath, result) {
        if (pageItem.typename === "PlacedItem" && pageItem.file && pageItem.file.fsName == decodedPath) {
            result.push(pageItem);
            return;
        }
        if (pageItem.pageItems && pageItem.pageItems.length > 0) {
            for (var i = 0; i < pageItem.pageItems.length; i++) {
                collectPlacedItems(pageItem.pageItems[i], decodedPath, result);
            }
        }
    }

    // 配置画像のバウンディングボックスと行列から長辺・短辺(pt)を計算する関数
    function getPlacedItemLongShortSides(placedItem) {
        var bbox = placedItem.boundingBox;
        var m = placedItem.matrix;
        var origWidthPt  = Math.abs(bbox[3] - bbox[1]);
        var origHeightPt = Math.abs(bbox[2] - bbox[0]);
        var scaleX = Math.sqrt(m.mValueA * m.mValueA + m.mValueB * m.mValueB);
        var scaleY = Math.sqrt(m.mValueC * m.mValueC + m.mValueD * m.mValueD);
        var widthPt  = origWidthPt  * scaleX;
        var heightPt = origHeightPt * scaleY;
        return { widthPt: widthPt, heightPt: heightPt };
    }

    // ドキュメント全体から該当画像の配置アイテムを収集
    var doc = app.activeDocument;
    var foundItems = [];
    for (var i = 0; i < doc.pageItems.length; i++) {
        collectPlacedItems(doc.pageItems[i], decodedPath, foundItems);
    }
    if (!foundItems.length) return null;

    // 最大面積の配置画像を選択し、そのサイズとリンク状態を取得
    var maxArea = 0;
    var maxSize = null;
    var maxLinkStatus = 0; // リンク状態初期値（正常）
    for (var j = 0; j < foundItems.length; j++) {
        var size = getPlacedItemLongShortSides(foundItems[j]);
        var area = size.widthPt * size.heightPt;
        if (area > maxArea) {
            maxArea = area;
            maxSize = size;
            if (typeof foundItems[j].linkStatus !== "undefined") {
                maxLinkStatus = foundItems[j].linkStatus;
            } else {
                maxLinkStatus = 0;
            }
        }
    }

    if (!maxSize) return null;

    // 長辺を算出しJSON文字列で返す
    var longSidePt = Math.max(maxSize.widthPt, maxSize.heightPt);

    return JSON.stringify({ longSidePt: longSidePt, linkStatus: maxLinkStatus });
}

// スクリプト実行開始
main();