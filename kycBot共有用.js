/*
* ver 1.3.1
* Last Updated on Mar. 12th 2020
*
* Reference:
* Library: M3W5Ut3Q39AaIwLquryEPMwV62A3znfOO
* API: Kokuzeicho API, Slack API, Slack Events API
* 下記URLよりコードを引用させていただいています：
* https://officeforest.org/wp/2019/04/23/google-apps-script%E3%81%A7xml%E3%82%92%E3%82%88%E3%81%97%E3%81%AA%E3%81%AB%E6%89%B1%E3%81%86%E6%96%B9%E6%B3%95/
*
*/

var importantInfo = "";
var tempText;
var incomingWebhookUrl = "https://hooks.slack.com/services/hagehage/higehige/hugehuge";


function testURLValidation(){
    var accessOptions = {
        "muteHttpExceptions" : true,
        "validateHttpsCertificates" : false,
        "followRedirects" : false,
    }
    var urlTest = "acoma-medical.co.jp";
    var xmlReqTest;

    if(IsError(UrlFetchApp.fetch(urlTest, accessOptions)) == 1){
        Logger.log(1);
    } else{
        Logger.log(2);
    }
}


function doGet(e){
    doPost(e);
    return
}

function doPost(e){
    var params = JSON.parse(e.postData.getDataAsString());

    var res = {};
    if(params.type == 'url_verification') {
        res = {'challenge':postData.challenge}
        return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
    }

    var channel = params.event.channel;
    var ts = params.event.ts;

    //Preventing identical messages from being sent spontaneously
    var cache = CacheService.getScriptCache();
    var cacheKey = channel + ':' + ts;
    var cached = cache.get(cacheKey);
    if (cached != null) {
        console.log('do nothing!');
        return;
    }
    cache.put(cacheKey, true, 600);

    //Slackのトリガーとして、一意なメッセージを指定する。
    const viablePattern = "法人が新規登録を行いました";
    var finalValidation;
    var searchTargets;
    var companyInfoArray;  
    var kycResearchArray;
    var bigAccountChecker = false;
    var mapUrl;

    if (!(params.type == "event_callback" && params.event.type == "message")) {
        return
    }
    if(!params.event.text.match(viablePattern)){
        return
    }

    searchTargets = arrangeResearchArray(params.event.text);

    if (isBigAccount(searchTargets[0]) === true ){
        bigAccountChecker = true;
    }

    if(bigAccountChecker){
        companyInfoArray[0] = "省略"; //ホームページ情報
        companyInfoArray[1] = "省略"; //業種情報
        kycResearchArray[0] = "省略"; //登記社名
        kycResearchArray[1] = "省略"; //登記住所
        kycResearchArray[2] = "省略"; //登記番号
        mapUrl = "省略";
    } else {
        companyInfoArray = getWebpageInfo(searchTargets[2]);
        mapUrl = getMapUrl(searchTargets[1]);
        kycResearchArray =  useKokuzeiAPI(searchTargets[0]);
        //memo: searchTargets[1]を使っていないのはバグではない。将来的に同一住所登録を検知するため。
    }

    if(importantInfo === ""){
        tempText = addToImportantInfo("なし");
    }
    finalValidation = importantInfo + companyInfoArray + kycResearchArray + mapUrl + params.event.ts;

    if(finalValidation.toString().match(viablePattern)){
        //最初に設定した「一意なメッセージ(viablePattern」がこの送信直前で含まれているか確認している。含まれている場合、無限ループになってしまわないようにエラーを出して送信処理を行わない。
        console.log("An error detected");
        return
    } else {
        return postMessage("Ver:1.3.0",importantInfo,companyInfoArray, kycResearchArray, mapUrl, params.event.ts);
    }

}


function arrangeResearchArray(text){

    var tempCompany;
    var tempAddress;
    var tempEmailAddress;
    var researchArray = ["", "", ""]; //社名、住所、メールアドレス

    if(!text){
        return researchArray;
    }

    tempCompany = text.match(/会社名:.*?\n/).toString();

    tempCompany = tempCompany.toString().replace(/会社名:/,"");
    tempCompany = tempCompany.toString().replace(/\n/,"");

    tempAddress = text.toString().match(/住所:.*?\n/);
    tempAddress = tempAddress.toString().replace(/住所:/,"");
    tempAddress = tempAddress.toString().replace(/\n/,"");

    tempEmailAddress = text.toString().match(/メールアドレス:.*?\n/);
    tempEmailAddress = tempEmailAddress.toString().replace(/メールアドレス: /,"");
    tempEmailAddress = tempEmailAddress.toString().replace(/ |(\n)/g,"");
    tempEmailAddress = tempEmailAddress.toString().replace(/<mailto:|>/g,"");

    researchArray[0] = tempCompany;
    researchArray[1] = tempAddress;
    researchArray[2] = tempEmailAddress;

    return researchArray;
}

function isBigAccount(companyName){

    //社内で担当者がいる会社
    const bigCompaniesPattern = /A社|B社|C社|D社|E社/;

    if(!companyName){
        return false
    }

    if(companyName.toString().match(bigCompaniesPattern)){
        tempText = addToImportantInfo('フィールドセールス');
        return true
    } else {
        return false
    }

}

function getWebpageInfo(emailAddress){

    const validationPattern = /outlook|yahoo|outlook|softbank|docomo|icloud|gmail|ocn|voda|au\.com|ezweb/;
    const filtering = /<meta([. ^\n\r])*?name="keywords".*?>|<meta([. ^\n\r])*?content="[. ]*?"[. ^\n\r]*?name="keywords"/;
    const metaKeywords = /<meta([. ^\n\r])*?name="keywords".*?>|<meta([. ^\n\r])*?content="[. ]*?"[. ^\n\r]*?name="keywords"/;
    const metaDescription = /<meta([. ^\n\r])*?name="description".*?>|<meta([. ^\n\r])*?content="[. ]*?"[. ^\n\r]*?name="description"/;

    var webpageArray = ['', '']; //返り値：websiteUrl, industryの格納用。初期化。
    var metaContents;
    var urlCandidate;
    var xmlrequest
    var accessOptions = {
        "muteHttpExceptions" : true,
        "validateHttpsCertificates" : false,
        "followRedirects" : false,
    };

    if(!emailAddress){
        webpageArray[0] = 'ホームページ確認できず。';
        webpageArray[1] = '要確認';
        return webpageArray;
    }

    if(emailAddress.toString().match(validationPattern)){
        webpageArray[0] = 'フリメと思われるのでホームページ要確認。';
        webpageArray[1] = '要確認';
        tempText = addToImportantInfo('フリーメール登録');
        return webpageArray;
    }

    urlCandidate = emailAddress.toString().replace(/.*@/,"")

    xmlrequest = UrlFetchApp.fetch(urlCandidate, accessOptions);

    switch (xmlrequest.getResponseCode()) {
        case(404):
            webpageArray[0] = 'webサイト確認できず(404)。';
            webpageArray[1] = '要確認';
            tempText = addToImportantInfo("404");
            break;
        case(403):
            webpageArray[0] = 'webサイト確認できず(403)。';
            webpageArray[1] = '要確認';
            tempText = addToImportantInfo("403");
            break;
        case(401):
            webpageArray[0] = 'webサイト確認できず(401)。';
            webpageArray[1] = '要確認';
            tempText = addToImportantInfo("401");
            break;
        case(400):
            console.log("400エラー：不正リクエスト");
            webpageArray[0] = 'webサイト不明(400)。';
            webpageArray[1] = '要確認';
            tempText = addToImportantInfo("400");
            break;
        default:
            metaContents = UrlFetchApp.fetch(urlCandidate).getContentText().toString();
            webpageArray[0] = urlCandidate;
            if (metaContents.toString().match(metaKeywords)) {
                metaContents = metaContents.toString().match(metaKeywords);
                metaContents = metaContents.toString().match(/content=".*?"/);
                metaContents = metaContents.toString().replace(/content=|"/g, "");
                webpageArray[1] = metaContents;
            }

            if(metaContents.toString().match(metaDescription)){
                metaContents = metaContents.toString().match(metaDescription);
                metaContents = metaContents.toString().match(/content=".*?"/);
                metaContents = metaContents.toString().replace(/content=|"/g, "");
                webpageArray[1] = metaContents;
            }

            if(!webpageArray[1]){
                webpageArray[1] = '取得できず。';
                break;
            }


            if(webpageArray[1].toString().match(/競合キーワード/)){
                /*
                * 上記の「競合キーワード」には、実際は自社が競合と捉えている業界のキーワードを指定する。
                * 例：自分が出前サービスを提供しているなら「フードデリバリー」など。
                */
                tempText = addToImportantInfo("競合");
                break;
            }

    }
    return webpageArray;
}

function useKokuzeiAPI(companyName){

    const validPattern = /合同会社|株式会社|有限会社|（合）|（株）|(有)/;
    const maekabuPattern = /^(合同会社|株式会社|有限会社|\（合\）|\（株\）|\(有\))/;
    const atokabuPatter = /(合同会社|株式会社|有限会社|\（合\）|\（株\）|\(有\))[ .　]*$/;

    const URLprefix = 'https://api.houjin-bangou.nta.go.jp/4/name?id=hogehoge&name=';
    const URLsuffix = '&type=12';

    var resultNumber = 0;
    var kokuzeiCompanyName = '';
    var kokuzeiCompanyNumber = '';
    var kokuzeiCompanyAddress = '';
    var prefecture = '';
    var city = '';
    var street = '';
    var response;
    var requestURL;

    var kokuzeiSearchedArray = [,,] //初期化
    if(!companyName){
        kokuzeiSearchedArray[0] = '社名検索エラー';
        kokuzeiSearchedArray[1] = '';
        kokuzeiSearchedArray[2] = '';
        return kokuzeiSearchedArray;
    }

    companyName = companyName.toString();


    if (companyName.toString().match(validPattern)){
        companyName = companyName.toString().replace(validPattern,"");
        companyName = zenkakuTransition(companyName);
        companyName = encodeURI(companyName);
    } else {
        kokuzeiSearchedArray[0] = "会社種不明のため検索せず"
        kokuzeiSearchedArray[1] = '';
        kokuzeiSearchedArray[2] = '';
        tempText = addToImportantInfo("会社種不明");
        return kokuzeiSearchedArray;
    }
    requestURL = URLprefix + companyName +URLsuffix;
    response = xmlToJson(UrlFetchApp.fetch(requestURL));

    resultNumber = response.corporations.count.Text;
    switch(resultNumber){
        case('1'): //1件なら表示すれば良い
            kokuzeiCompanyName = response.corporations.corporation.name.Text;
            kokuzeiCompanyNumber = response.corporations.corporation.corporateNumber.Text;

            prefecture =  response.corporations.corporation.prefectureName.Text;
            city = response.corporations.corporation.cityName.Text;
            street = response.corporations.corporation.streetNumber.Text;
            kokuzeiCompanyAddress = prefecture + city + street;

            kokuzeiSearchedArray[0] = kokuzeiCompanyName;
            kokuzeiSearchedArray[1] = "\n登記住所: " + kokuzeiCompanyAddress;
            kokuzeiSearchedArray[2] = "\n登記番号: " + kokuzeiCompanyNumber;
            break;
        case('0'):
            console.log("国税庁のAPIで検索結果は0件。");
            kokuzeiSearchedArray[0] = '登記簿に該当社名データなし';
            kokuzeiSearchedArray[1] = '';
            kokuzeiSearchedArray[2] = '';
            tempText = addToImportantInfo("登記簿登録なし");
            break;
        default:
            kokuzeiSearchedArray[0] = "登記簿に同社名で" + resultNumber + "件登録あり \n" ;
            kokuzeiSearchedArray[1] = '';
            kokuzeiSearchedArray[2] = '';
            tempText = addToImportantInfo("登記情報を特定できず");
            break;
    }
    return kokuzeiSearchedArray;
}

//検索のためにアルファベットを全角に変換(例:ABCdefはＡＢＣｄｅｆに変換される)
function zenkakuTransition(name){
    console.log("zenkakuTransition initiated.")
    if(!name){
        console.log("nameが空欄なため全角化を終了")
        name = "";
        return name
    }
    // 株式会社とかはついていないもののみ引数として受けとる
    name = name.toString();
    var pureName = name.replace(/[A-Za-z0-9]/g, function(s) {
        return String.fromCharCode(s.charCodeAt(0) + 65248);
    });
    Logger.log(pureName);
    return pureName;
}




function getMapUrl(address){
    var mapSearchUrl = "https://maps.google.co.jp/maps?q=";
    if(!address){
        console.log("No address typed!");
        return "要確認です"
    }
    return mapSearchUrl + address.toString().replace(/ |　/g,"");
}

function postMessage(version,info,companyInfoArray, kycArray, mapUrl, time) {

    var url = incomingWebhookUrl;

    var payload = {
        "text": "test Body",
        "blocks": [
            {
                "type": "section",
                "text": {
                    "type": "mrkdwn",
                    "text": info
                }
            },
            {
                "type": "section",
                "text": {
                    "type": "mrkdwn",
                    "text": ">法人登記情報: \n" + kycArray[0] +  kycArray[1] +  kycArray[2]
                }
            },

            {
                "type": "section",
                "text": {
                    "type": "mrkdwn",
                    "text": ">検索リンク: \n <"+ mapUrl + "|GoogleMaps>" + "、 <https://www." + companyInfoArray[0] + "|ホームページ> \n "+ version
                }
            },
        ]
    };

    //make the message as a reply when a time parameter is given
    if(time){
        payload.thread_ts = time;
    }

    var options = {
        "method" : "post",
        "contentType" : "application/x-www-form-urlencoded",
        "text":"text message here",
        "payload" : JSON.stringify(payload)
    };

    UrlFetchApp.fetch(url, options);
    return

}


function addToImportantInfo(text){
    if(!text){
        return
    }
    switch(text){
        case("フリーメール登録"):
            importantInfo = importantInfo + ":kyc_free_email: ";
            break;
        case("競合"):
            importantInfo = importantInfo + ":kyc_kyougou:";
            break;
        case("会社種不明"):
            importantInfo = importantInfo + ":kyc_unknown:(会社種) ";
            break;
        case("登記情報を特定できず"):
            importantInfo = importantInfo + ":kyc_unknown:(登記情報) ";
            break;
        case("登記簿記録なし"):
            importantInfo = importantInfo + ":kyc_unknown:(登記情報) ";
            break;
        case("フィールドセールス"):
            importantInfo = importantInfo + ":kyc_fs: ";
            break;
        case("同住所登録あり"):
            importantInfo = importantInfo + ":kyc_sameAddress:→同住所での登録があるようです。確認をお願いします。\n";
            break;
        case("404"):
            importantInfo = importantInfo + ":kyc_unknown:(メルアド) ";
            break;
        case("403"):
            importantInfo = importantInfo + ":kyc_unknown:(メルアド) ";
            break;
        case("401"):
            importantInfo = importantInfo + ":kyc_unknown:(メルアド) ";
            break;
        case("400"):
            importantInfo = importantInfo + ":kyc_unknown:(メルアド) ";
            break;
        case("なし"):
            importantInfo = importantInfo + "特記事項なし";
            break;
        default:
            break;
    }
    return
}


//以下、このURLから引用させていただきました。
//https://officeforest.org/wp/2019/04/23/google-apps-script%E3%81%A7xml%E3%82%92%E3%82%88%E3%81%97%E3%81%AA%E3%81%AB%E6%89%B1%E3%81%86%E6%96%B9%E6%B3%95/
//XML2JSONで処理をする
function xmlconvert(){
    console.log("xmlconvert initiated.")

    //検索結果を取得する
    var response = UrlFetchApp.fetch(sURL);

    //JSONデータに変換する
    var jsondoc = xmlToJson(response);

    //結果を出力する
    Logger.log(jsondoc);

}

//XMLをJSONに変換するとき利用する関数
function xmlToJson(xml) {
    console.log("xmlToJson initiated.")
    //XMLをパースして変換関数に引き渡し結果を取得する
    var doc = XmlService.parse(xml);
    var result = {};
    var root = doc.getRootElement();
    result[root.getName()] = elementToJson(root);
    return result;
}

//変換するメインルーチン
function elementToJson(element) {
    console.log("elementToJson initiated.")

    //結果を格納する箱を用意
    var result = {};

    // Attributesを取得する
    element.getAttributes().forEach(function(attribute) {
        result[attribute.getName()] = attribute.getValue();
    });

    //Child Elementを取得する
    element.getChildren().forEach(function(child) {
        //キーを取得する
        var key = child.getName();

        //再帰的にもう一度この関数を実行して判定
        var value = elementToJson(child);

        //XMLをJSONに変換する
        if (result[key]) {
            if (!(result[key] instanceof Array)) {
                result[key] = [result[key]];
            }
            result[key].push(value);
        } else {
            result[key] = value;
        }
    });

    //タグ内のテキストデータを取得する
    if (element.getText()) {
        result['Text'] = element.getText();
    }
    return result;
}