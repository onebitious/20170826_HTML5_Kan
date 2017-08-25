var doc = app.activeDocument; //ドキュメント
var sel = doc.selection; //選択している文字
var embedJS = '<script src="https://use.typekit.net/wxp7vtm.js"></script><script>try{Typekit.load({ async: true });}catch(e){}</script>'; //TypeKitの埋め込みコード

var outPutHTML = getContents();

/*////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
タグ、セレクタ取得関数定義
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////*/
function getContents() {
    //▼選択している文字を取得
    if (sel.typename == "TextRange") {
        var myContents = sel.contents;
    } else if (sel.length != 0) {
        alert("フレームが選択されています。" + "\r\n" + "文字を選択してください。");
        return;
    } else {
        alert("文字が選択されていません。" + "\r\n" + "文字を選択してください。");
        return;
    }

    //▼タグ付与のダイアログ表示
    var tagFlag = false; //フラグの初期化
    var textFlag = true;
    var myWindow = new Window('dialog', '要素・属性の設定', [800, 500, 1025, 730]);
    myWindow.elementText = myWindow.add("statictext", [10, 10, 450, 30], "要素を選択してください。"); //固定テキスト
    myWindow.tagSelectList = myWindow.add("dropdownlist", [10, 35, 200, 55], ["h1", "h2", "h3", "h4", "h5", "h6", "p"]);
    myWindow.tagSelectList.selection = 6; //デフォルト表示は一番下のもの
    //myWindow.pBottom = myWindow.add("radiobutton", [105, 35, 205, 65], "p タグ");
    myWindow.attributeText = myWindow.add("statictext", [10, 70, 275, 100], "属性名を選択してください"); //固定テキスト
    myWindow.dropdownList = myWindow.add("dropdownlist", [10, 100, 200, 120], ["id セレクタ", "class セレクタ", ""]);
    myWindow.dropdownList.selection = 0; //デフォルト表示は一番上のもの
    myWindow.attributeValueText = myWindow.add("statictext", [10, 150, 275, 170], "属性値を入力してください"); //固定テキスト

    myWindow.editText = myWindow.add("edittext", [10, 170, 200, 190], "", {
        readonly: false
    }); //id、classセレクタだったら入力できる
    myWindow.dropdownList.onChange = function () {
        if (myWindow.dropdownList.selection == 2) {
            myWindow.editText = myWindow.add("edittext", [10, 170, 200, 190], "※属性名空欄時は入力不可", {
                readonly: true
            }); //入力できなくする
            textFlag = false; //フラグ
        } else {
            myWindow.editText = myWindow.add("edittext", [10, 170, 200, 190], "", {
                readonly: false
            }); //id、classセレクタだったら入力できる
        }
    }

    myWindow.okBottom = myWindow.add("button", [20, 200, 100, 220], "OK", {
        name: "ok"
    });
    myWindow.cancelBottom = myWindow.add("button", [120, 200, 200, 220], "キャンセル", {
        name: "cancel"
    });

    var bottomFlag = myWindow.show(); //ダイアログを表示し、OK、キャンセルボタンの結果を取得

    if (bottomFlag == 2) { //キャンセルの場合処理を抜ける
        return;
    }
    var tagSelectListResult = myWindow.tagSelectList.selection.text;; //ドロップダウンリストから選択したタグ名が返る
    var dropDownListResult = myWindow.dropdownList.selection.text; //ドロップダウンリストから選択したセレクタ名が返る

    var mySelectorName = myWindow.editText.text; //入力したテキストが返る

    //▼セレクタ名の取得。属性値が空の場合の処理
    if (textFlag == false) {
        var selectorNameReault = "";
    } else if (mySelectorName == "") {
        var selectorNameReault = '="○○○○●"';
        var mySelectorName = '○○○○●';
    } else {
        var selectorNameReault = '="' + mySelectorName + '"';
    }

    //▼セレクタの選択
    if (dropDownListResult == "id セレクタ") {
        var selectorResult = "id";
        var chiceSelector = "#" + mySelectorName;
    } else if (dropDownListResult == "class セレクタ") {
        var selectorResult = "class";
        var chiceSelector = "." + mySelectorName;

    } else {
        var selectorResult = "";
    }

    //▼タグ
    var tagResult = "<" + tagSelectListResult + " " + selectorResult + selectorNameReault + ">" + myContents + "</" + tagSelectListResult + ">";

    /*////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    OpenType機能
    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////*/
    //▼OpenType数字
    switch (sel.figureStyle) {
        case FigureStyleType.DEFAULTFIGURESTYLE:
            figStyCSS = ""; //デフォルトの数字
            break;
        case FigureStyleType.TABULAR:
            figStyCSS = ""; //等幅ライニング数字
            break;
        case FigureStyleType.PROPORTIONALOLDSTYLE:
            figStyCSS = 'font-feature-settings: "onum";'; //オールドスタイル数字
            break;
        case FigureStyleType.PROPORTIONAL:
            figStyCSS = 'font-feature-settings: "lnum";'; //ライニング数字
            break;
        case FigureStyleType.TABULAROLDSTYLE:
            figStyCSS = ""; //等幅オールドスタイル数字
            break;
    }

    //▼OpenType位置
    switch (sel.openTypePosition) {
        case FontOpenTypePositionOption.OPENTYPEDEFAULT:
            figStyCSS = ''; //デフォルトの位置
            break;
        case FontOpenTypePositionOption.OPENTYPESUPERSCRIPT:
            figStyCSS = 'font - feature - settings: "sups";'; //上付き文字
            break;
        case FontOpenTypePositionOption.OPENTYPESUBSCRIPT:
            figStyCSS = 'font-feature-settings: "subs";'; //下付き文字
            break;
        case FontOpenTypePositionOption.NUMERATOR:
            figStyCSS = ''; //分子
            break;
        case FontOpenTypePositionOption.DENOMINATOR:
            figStyCSS = ''; //分母
            break;
    }

    //▼OpenTypeその他
    if (sel.ligature == true) {
        var ligaResult = 'font-feature-settings: "liga", "clig";'; //欧文合字
    } else {
        var ligaResult = "";
    }
    if (sel.ordinals == true) {
        var ordnResult = 'font-feature-settings: "ordn";'; //上付き序文表記
    } else {
        var ordnResult = "";
    }
    if (sel.contextualLigature == true) {
        var caltResult = 'font-feature-settings: "calt";'; //前後関係に依存する文字
    } else {
        var caltResult = "";
    }
    if (sel.fractions == true) {
        var fracResult = 'font-feature-settings: "frac";'; //スラッシュを用いた分数
    } else {
        var fracResult = "";
    }
    if (sel.discretionaryLigature == true) {
        vardligResult = 'font-feature-settings: "dlig";'; //任意の合字
    } else {
        var vardligResult = "";
    }
    if (sel.proportionalMetrics == true) {
        var paltResult = 'font-feature-settings: "palt";'; //プロポーショナルメトリックス
    } else {
        var paltResult = "";
    }
    if (sel.swash == true) {
        var swshResult = 'font-feature-settings: "swsh";'; //スワッシュ字形
    } else {
        var swshResult = "";
    }
    if (sel.stylisticAlternates == true) {
        var saltResult = 'font-feature-settings: "salt";'; //デザインのバリエーション
    } else {
        var saltResult = "";
    }
    if (sel.italics == true) {
        var italResult = 'font-feature-settings: "ital";'; //欧文イタリック
    } else {
        var italResult = "";
    }

    //▼セレクタに書き出す
    var myContents = sel.contents; //選択している文字
    var myFontName = sel.textFont.name; //フォント名
    //▼webフォント名取得
    switch (myFontName) {
        case "WarnockPro-Regular":
            myFontName = "warnock-pro";
            break;
        case "BickhamScriptPro3-Bold":
            myFontName = "bickham-script-pro-3";
            break;
        case "Bree-Regular":
            myFontName = "bree";
            break;
        case "SourceHanSerif-Regular":
            myFontName = "source-han-serif-japanese";
            break;
        case "KozMinPro-Regular":
            myFontName = "kozuka-mincho-pro"; //小塚明朝Pro R
            break;
        case "FutoGoB101Pr6-Bold":
            myFontName = "a-otf-futo-go-b101-pr6n"; //太ゴ
            break;
        case "MidashiGoPr6N-MB31":
            myFontName = "a-otf-midashi-go-mb31-pr6n"; //見出しゴ
            break;
        case "GothicBBBPr6N-Medium":
            myFontName = "a-otf-gothic-bbb-pr6n"; //中ゴ
            break;
    }
    var myFontStyle = sel.textFont.style; //スタイル名
    //▼RGB値を取得する関数
    function getColorValue() {
        var myFontColorRed = sel.fillColor.red; //red値を取得
        var myFontColorGreen = sel.fillColor.green; //green値を取得
        var myFontColorBlue = sel.fillColor.blue; //blue値を取得
        var myColor = myFontColorRed + ',' + myFontColorGreen + ',' + myFontColorBlue;
        return myColor;
    }
    var myFontSize = sel.size + 'pt'; //サイズ
    //▼CSSに書き出す
    if (textFlag == false) {
        //▼セレクタがない場合
        var selectorResult = tagSelectListResult + '{' + '\r\n' + 'font-family: "' + myFontName + '";' + '\r\n' + 'font-size: ' + myFontSize + ';' + '\r\n' + 'color:rgb(' + getColorValue() + ');' + '\r\n' + figStyCSS + figStyCSS + figStyCSS + figStyCSS + figStyCSS + figStyCSS + figStyCSS + figStyCSS + figStyCSS + figStyCSS + ligaResult + ordnResult + caltResult + fracResult + vardligResult + paltResult + swshResult + saltResult + italResult + '}';
    } else {
        //▼セレクタがある場合       
        var selectorResult = chiceSelector + '{' + '\r\n' + 'font-family: "' + myFontName + '";' + '\r\n' + 'font-size: ' + myFontSize + ';' + '\r\n' + 'color:rgb(' + getColorValue() + ');' + '\r\n' + figStyCSS + figStyCSS + figStyCSS + figStyCSS + figStyCSS + figStyCSS + figStyCSS + figStyCSS + figStyCSS + figStyCSS + ligaResult + ordnResult + caltResult + fracResult + vardligResult + paltResult + swshResult + saltResult + italResult + '}';
    }
    var htmlTag = '<!DOCTYPE html><html lang="ja"><head>' + embedJS + '<meta charset="UTF-8"><title>Document</title><style>' + selectorResult + '</style></head><body>' + tagResult + '</body></html>';
    return htmlTag;
}

/*////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
ファイル保存
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////*/
if (outPutHTML !== undefined) {
    saveFile();
}

function saveFile() {
    var inputFileName = File.saveDialog("保存先とファイル名を指定してください。");
    if (inputFileName == null) {
        //alert("キャンセルしました。");
        return;
    }

    var inputFileName = inputFileName.toString();
    var myNewText = inputFileName.match(/.html/);
    if (myNewText) {
        var myNewText = inputFileName;
    } else {
        var inputFileName = inputFileName + ".html";
    }

    var fileObj = new File(inputFileName);
    var inputFileName = fileObj.open("w");
    if (inputFileName == true) {
        fileObj.encoding = "UTF-8";
        fileObj.writeln(outPutHTML);
        fileObj.close;
        alert("処理が終わりました。");
    }
}
