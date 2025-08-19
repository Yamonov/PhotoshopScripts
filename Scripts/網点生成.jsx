/*
<javascriptresource>
<name>疑似網点生成CMYK</name>
<category>YPresets</category>
</javascriptresource>
*/

(function () {

var isGray = app.activeDocument.mode === DocumentMode.GRAYSCALE;
var presetTable = [
  ["CMYK", "K45_175線", 15, 75, 0, 45, 175, 175, 175, 175],
  ["CMYK", "M45_175線", 15, 45, 0, 75, 175, 175, 175, 175],
  ["CMYK", "K45_80線", 15, 75, 0, 45, 80, 80, 80, 80],
  ["CMYK", "K45_300線", 15, 75, 0, 45, 300, 300, 300, 300],
  ["CMYK", "7度_175（Y190）線", 22, 52, 7, 82, 175, 175, 190, 175],
  ["GRAY", "45度_175線", 45, 175],
  ["GRAY", "45度_80線", 45, 80],
  ["GRAY", "45度_300線", 45, 300],
  ["GRAY", "75度_175線", 75, 175],
  ["GRAY", "0度_175線", 0, 175]
];
var keys = isGray ? ["Gray"] : ["C", "M", "Y", "K"];

function buildDialog() {
  var dlg = new Window("dialog", "線数と角度の設定");
  dlg.orientation = "column";
  dlg.alignChildren = "left";

  var mode = isGray ? "GRAY" : "CMYK";
  var filtered = [];
  for (var i = 0; i < presetTable.length; i++) {
    if (presetTable[i][0] === mode) {
      filtered.push(presetTable[i]);
    }
  }
  var presetNames = [];
  for (var i = 0; i < filtered.length; i++) {
    presetNames.push(filtered[i][1]);
  }

  var presetList = dlg.add("dropdownlist", undefined, presetNames);
  presetList.selection = 0;

  var inputGroup = dlg.add("group");
  inputGroup.orientation = "column";
  var inputFields = {};

  for (var i = 0; i < keys.length; i++) {
    var row = inputGroup.add("group");
    row.add("statictext", undefined, keys[i] + ": 角度");
    inputFields[keys[i]] = {};
    inputFields[keys[i]].angle = row.add("edittext", undefined, "0");
    inputFields[keys[i]].angle.characters = 5;
    row.add("statictext", undefined, "線数");
    inputFields[keys[i]].lpi = row.add("edittext", undefined, "0");
    inputFields[keys[i]].lpi.characters = 5;
  }

  var resGroup = dlg.add("group");
  resGroup.add("statictext", undefined, "出力解像度:");
  var resDropdown = resGroup.add("dropdownlist", undefined, ["4800", "2400", "1200", "600"]);
  resDropdown.selection = 1;
  inputFields.resolution = resDropdown;

  presetList.onChange = function () {
    var sel = filtered[presetList.selection.index];
    if (mode === "GRAY") {
      inputFields["Gray"].angle.text = sel[2];
      inputFields["Gray"].lpi.text = sel[3];
    } else {
      var angles = sel.slice(2, 6);
      var lpis = sel.slice(6);
      var ch = ["C", "M", "Y", "K"];
      for (var i = 0; i < ch.length; i++) {
        inputFields[ch[i]].angle.text = angles[i];
        inputFields[ch[i]].lpi.text = lpis[i];
      }
    }
  };
  presetList.notify();
  // 注意文をパネルで追加
  var notePanel = dlg.add("panel", undefined, "注意");
  notePanel.orientation = "column";
  notePanel.alignChildren = "left";
  notePanel.margins = [10, 15, 10, 10];

  var note = notePanel.add("statictext", undefined,
    "元画像を複製して処理します。\n※先に画像解像度で、使用サイズ＆解像度にリサイズし、統合してから実行してください。\nESCでキャンセル",
    { multiline: true }
  );
  note.maximumSize.width = 380;
  dlg.add("group").add("button", undefined, "OK");

  return dlg.show() == 1 ? inputFields : null;
}

function getScreenSettings(fields) {
  var s = {};
  for (var i = 0; i < keys.length; i++) {
    s[keys[i]] = {
      angle: parseFloat(fields[keys[i]].angle.text),
      lpi: parseFloat(fields[keys[i]].lpi.text)
    };
  }
  s.resolution = parseInt(fields.resolution.selection.text, 10);
  return s;
}

function binarizeChannel(doc, setting, resolution) {
  app.activeDocument = doc;
  // doc.resizeImage(undefined, undefined, resolution, ResampleMethod.NONE); // 解像度変更を無効化
  var bmo = new BitmapConversionOptions();
  bmo.resolution = resolution;
  bmo.method = BitmapConversionType.HALFTONESCREEN;
  bmo.frequency = setting.lpi;
  bmo.angle = setting.angle;
  bmo.shape = BitmapHalfToneType.ROUND;
  doc.changeMode(ChangeMode.BITMAP, bmo);
  doc.changeMode(ChangeMode.GRAYSCALE);
}

function mergeCMYKChannels(docNames) {
  var desc = new ActionDescriptor();
  var list = new ActionList();
  for (var i = 0; i < 4; i++) {
    var ref = new ActionReference();
    ref.putName(stringIDToTypeID("document"), docNames[i]);
    list.putReference(ref);
  }
  desc.putList(stringIDToTypeID("null"), list);
  desc.putEnumerated(stringIDToTypeID("mode"), stringIDToTypeID("colorSpace"), stringIDToTypeID("CMYKColorEnum"));
  try {
    executeAction(stringIDToTypeID("mergeChannels"), desc, DialogModes.NO);
  } catch (e) {
    alert("mergeChannels に失敗しました\n" + e);
  }
}

if (app.documents.length === 0 ||
    !(app.activeDocument.mode === DocumentMode.CMYK || app.activeDocument.mode === DocumentMode.GRAYSCALE)) {
  alert("CMYKまたはグレースケールモードの画像を開いてから実行してください。");
  return;
}

var inputFields = buildDialog();
if (!inputFields) return;

var screenSettings = getScreenSettings(inputFields);
var resolution = screenSettings.resolution;
var doc = app.activeDocument;
var dup = doc.duplicate(doc.name.replace(/\.[^\.]+$/, "") + "_複製", true);
try {
  dup.flatten();
} catch (e) {
  // 統合済みで flatten が不要な場合はスキップ
}
app.activeDocument = dup;
var splitDocs;
if (isGray) {
  splitDocs = [dup];
} else {
  splitDocs = dup.splitChannels();
}

var binarizedDocs = [];
for (var i = 0; i < keys.length; i++) {
  binarizeChannel(splitDocs[i], screenSettings[keys[i]], resolution);
  binarizedDocs.push(splitDocs[i]);
}

if (!isGray) {
  var docNames = [];
  for (var i = 0; i < binarizedDocs.length; i++) {
    docNames.push(binarizedDocs[i].name);
  }
  mergeCMYKChannels(docNames);
}

})();
