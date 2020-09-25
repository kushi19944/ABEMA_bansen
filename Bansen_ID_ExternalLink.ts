import RPA from 'ts-rpa';
import {
  WebDriver,
  By,
  FileDetector,
  Key,
  WebElement,
} from 'selenium-webdriver';

var fs = require('fs');

// 読み込みする スプレッドシートID と シート名 の記載
const SSID = process.env.BANSEN_SheetID; //
const SSName1 = process.env.BANSEN_SheetName_External; //
// 画像などを保存するフォルダのパスを記載する。
const DownloadFolder = __dirname + '/Download/';
// Abematvのログイン。 メールアドレス・パスワードの記載 <<漏洩注意>>
const AbematvID = process.env.AbemaID;
const AbematvPW = process.env.AbemaPW;
// AAAMS 本番環境のログイン ID / PW
const AAAMS_ID = process.env.AAAMS_ID; //
const AAAMS_PW = process.env.AAAMS_PW; //
// SlackのトークンとチャンネルID
const SlackToken = process.env.AbemaTV_hubot_Token;
const SlackChannel = process.env.AbemaTV_Bansen_Channel;

const FirstLoginFlag = ['true'];

async function WorkStart() {
  await RPA.Google.authorize({
    //accessToken: process.env.GOOGLE_ACCESS_TOKEN,
    refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
    tokenType: 'Bearer',
    expiryDate: parseInt(process.env.GOOGLE_EXPIRY_DATE, 10),
  });
  // 作業用フォルダを空にするために削除する
  DeleteFiles();
  // アセットID と クリエイティブID を保持する変数
  const IDLIST = [['', '']];
  // 作業行う Row と データ を取得する
  const WorkData = [];
  const Row = [];
  await DataRowGet(WorkData, Row);
  RPA.Logger.info(WorkData[0]);
  // アセット名ごとに 【外部リンク付き番宣】/【買えるAbemaTV】でアカウントを切り替えるようにする
  const AccountFlag = ['null'];
  if (WorkData[0][0][3].indexOf('【外部リンク付き') >= 0) {
    AccountFlag[0] = '【外部リンク付き番宣】';
  }
  if (WorkData[0][0][3].indexOf('【買えるAbemaTV') >= 0) {
    AccountFlag[0] = '【買えるAbemaTV】';
  }
  // ヘッドレスモードでもGoogleDriveから動画ファイルダウンロード
  await HeadlessModeDownload(WorkData[0]);
  // AAAMS　アカウントにログインする
  if (FirstLoginFlag[0] == 'true') {
    await AAAMS_Login(AccountFlag);
  }
  if (FirstLoginFlag[0] == 'false') {
    await AAAMS_2nd_Login(AccountFlag);
  }
  // 一度ログインしたら、次はログインページをスキップさせる
  FirstLoginFlag[0] = 'false';
  // ローカルフォルダーから .mp4動画のファイルパスを取得する
  const FilePathData = [''];
  await FilePathGet(FilePathData);
  // AAAMSへファイルアップロード・トランスコード実行テスト
  await NewAsset_Upload_function(FilePathData);
  // トランスコード完了後に、　アセット入力する
  await NewAsset_Create_function(IDLIST, WorkData[0], Row);
  // アセットIDを取得する
  await Get_AssetID_function(IDLIST, WorkData[0], Row);
  if (AccountFlag[0] == '【外部リンク付き番宣】') {
    await RPA.WebBrowser.get(process.env.AAAMS_Account_1);
  }
  if (AccountFlag[0] == '【買えるAbemaTV】') {
    await RPA.WebBrowser.get(process.env.AAAMS_Account_2);
  }
  // クリエイティブ作成を行う
  await CreativeCreate(IDLIST, WorkData[0], Row);
  // クリエイティブIDを取得する
  await CreativeIDGet(IDLIST, WorkData[0], Row);
  // スプシにアセット・クリエイティブIDを貼り付ける
  await Spreadsheet_setValues_function(IDLIST, WorkData[0], Row);

  RPA.Logger.info(IDLIST);
  await RPA.sleep(1500);
}

async function Start() {
  await RPA.Google.authorize({
    //accessToken: process.env.GOOGLE_ACCESS_TOKEN,
    refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
    tokenType: 'Bearer',
    expiryDate: parseInt(process.env.GOOGLE_EXPIRY_DATE, 10),
  });
  const MainLoopFlag = [];
  MainLoopFlag[0] = false;
  while (0 == 0) {
    const AllData = await RPA.Google.Spreadsheet.getValues({
      spreadsheetId: `${SSID}`,
      range: `${SSName1}!A3:K3000`,
    });
    for (let i in AllData) {
      RPA.Logger.info(AllData[i]);
      if (String(AllData[i][3]).length < 1) {
        MainLoopFlag[0] = true;
        break;
      }
      if (AllData[i][10] == '初期') {
        if (AllData[i]) await WorkStart();
      }
    }
    if (MainLoopFlag[0] == true) {
      break;
    }
  }
  await RPA.WebBrowser.quit();
}

Start();

async function HeadlessModeDownload(AllData) {
  await RPA.sleep(2000);
  const DriveID = [];
  // 配列から DriveID だけを抽出する
  for (let i in AllData) {
    if (AllData[i][2].indexOf('https://drive.google.com/open?id=') == 0) {
      const data = AllData[i][2].split('https://drive.google.com/open?id=');
      DriveID.push(data[1]);
    }
    if (AllData[i][2].indexOf('https://drive.google.com/file/d/') == 0) {
      const data = AllData[i][2].split('https://drive.google.com/file/d/');
      const data2 = data[1].split('/');
      DriveID.push(data2[0]);
    }
  }
  RPA.Logger.info(DriveID);
  await RPA.Google.Drive.download({ fileId: `${DriveID}` });
}

// ダウンロードしたファイルを削除する関数
async function DeleteFiles() {
  const path = require('path');
  const firstdata = fs.readdirSync(DownloadFolder);
  RPA.Logger.info(firstdata);
  for (let i in firstdata) {
    fs.unlink(`${DownloadFolder}${firstdata[i]}`, function (err) {});
  }
}

async function AAAMS_Login(AccountFlag) {
  await RPA.WebBrowser.get(process.env.AAAMS_Login_URL);
  //await RPA.WebBrowser.get(`https://stg-admin.vega.fm/#/asset`);
  await RPA.sleep(2000);
  try {
    const AAAMS_loginID_ele = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({
        name: 'email',
      }),
      8000
    );
    await RPA.WebBrowser.sendKeys(AAAMS_loginID_ele, [AAAMS_ID]);
    const AAAMS_loginPW_ele = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({
        name: 'password',
      }),
      8000
    );
    await RPA.WebBrowser.sendKeys(AAAMS_loginPW_ele, [AAAMS_PW]);
    const AAAMS_LoginNextButton = await RPA.WebBrowser.findElementsByClassName(
      'auth0-lock-submit'
    );
    await RPA.WebBrowser.mouseClick(AAAMS_LoginNextButton[0]);
    await RPA.sleep(3000);
  } catch {
    RPA.Logger.info('ログイン画面飛ばします');
  }

  await RPA.sleep(2000);
  // チャンネル更新画面が出るため待機する
  try {
    const ChannelAlart = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({
        xpath: '/html/body/div/div/div[6]/div[2]/header/div',
      }),
      5000
    );
    const ChannelAlartButton = await RPA.WebBrowser.findElementByXPath(
      '/html/body/div/div/div[6]/div[2]/footer/div[2]'
    );
    const AlartText = await ChannelAlart.getText();
    if (AlartText == '下記更新されました。設定を確認してください。') {
      await RPA.WebBrowser.mouseClick(ChannelAlartButton);
      await RPA.sleep(1500);
    }
  } catch {}
  try {
    const Alart = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({
        xpath: '/html/body/div/div/div[5]/div[2]/div/p',
      }),
      3000
    );
    const Alartbutton = await RPA.WebBrowser.findElementByXPath(
      '/html/body/div/div/div[5]/div[2]/footer/div[1]'
    );
    await RPA.WebBrowser.mouseClick(Alartbutton);
    await RPA.sleep(2000);
  } catch {
    RPA.Logger.info('AAAMS アラート出ませんでしたので次に進みます');
  }
  if (AccountFlag == '【外部リンク付き番宣】') {
    RPA.Logger.info('外部リンク付き自社広告アカウント直接呼び出しします');
    await RPA.WebBrowser.get(process.env.AAAMS_Account_AssetURL_1);
  }
  if (AccountFlag == '【買えるAbemaTV】') {
    RPA.Logger.info('買えるAbemaTV社アカウント直接呼び出しします');
    await RPA.WebBrowser.get(process.env.AAAMS_Account_AssetURL_2);
  }
  await RPA.sleep(3000);
  // 変な更新画面が出るのでスルーする
  try {
    const Koushin = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({
        xpath: '/html/body/div[1]/div/div[5]/div[2]/header/div',
      }),
      5000
    );
    const KoushinText = await Koushin.getText();
    if (KoushinText.length > 1) {
      const NextButton01 = await RPA.WebBrowser.findElementByXPath(
        '/html/body/div[1]/div/div[5]/div[2]/footer/div[1]'
      );
      await RPA.WebBrowser.mouseClick(NextButton01);
      await RPA.sleep(1000);
    }
  } catch {
    RPA.Logger.info('謎の更新画面出ませんでした');
  }
}

async function AAAMS_2nd_Login(AccountFlag) {
  if (AccountFlag == '【外部リンク付き番宣】') {
    RPA.Logger.info('外部リンク付き自社広告アカウント直接呼び出しします');
    await RPA.WebBrowser.get(process.env.AAAMS_Account_AssetURL_1); //直接外部リンク用のアセットページへ飛ぶ
    await RPA.sleep(3000);
    /*
    await RPA.WebBrowser.get(
      `https://stg-admin.vega.fm/#/campaign/{%22account%22:{%22ids%22:[%22ac_6%22]},%22accountType%22:%22BANSEN_AD%22}`
    );
    */
  }
  if (AccountFlag == '【買えるAbemaTV】') {
    RPA.Logger.info('買えるAbemaTV社アカウント直接呼び出しします');
    await RPA.WebBrowser.get(process.env.AAAMS_Account_AssetURL_2);
    await RPA.sleep(3000);
    /*
    await RPA.WebBrowser.get(
      `https://stg-admin.vega.fm/#/campaign/{%22account%22:{%22ids%22:[%22ac_6%22]},%22accountType%22:%22BANSEN_AD%22}`
    );
    */
  }
  await RPA.sleep(2300);
}

async function DataRowGet(WorkData, Row) {
  const firstrow = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName1}!K3:K2000`,
  });
  for (let i in firstrow) {
    if (firstrow[i][0].indexOf('初期') == 0) {
      Row[0] = Number(i) + 3;
      break;
    }
  }
  RPA.Logger.info('この行の作業実行します → ', Row[0]);
  // ID円滑シート から作業する行のデータを抽出する
  WorkData[0] = await RPA.Google.Spreadsheet.getValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName1}!A${Row[0]}:K${Row[0]}`,
  });
  // ID円滑シート　のM列に作業中と記入する
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName1}!K${Row[0]}:K${Row[0]}`,
    values: [['作業中']],
  });
  RPA.Logger.info(`${Row[0]}　行目のステータスを 作業中 に変更しました`);
}

// フォルダーから動画のパスを取得する関数
async function FilePathGet(FilePathData) {
  const path = require('path');
  const dirPath = path.resolve(DownloadFolder);
  const firstdata = [];
  firstdata[0] = fs.readdirSync(dirPath);
  RPA.Logger.info(firstdata[0]);
  // .mp4　が含まれているファイルだけ抽出する
  for (let i in firstdata[0]) {
    if (firstdata[0][i].indexOf('.mp4') > 1) {
      FilePathData[0] = firstdata[0][i];
    }
  }
  RPA.Logger.info(FilePathData);
}

// クリエイティブを入力する関数
async function CreativeCreate(IDLIST, Datas, Row) {
  await RPA.sleep(2000);
  const CreativeCreateButton = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      className: 'Button__button___2uDT- Button__create___1F14z',
    }),
    5000
  );
  await CreativeCreateButton.click();
  await RPA.sleep(1000);
  // アセットを紐付ける
  await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      name: 'creativeAsset',
    }),
    5000
  );
  // アセット名で検索
  const AssetSearchInput: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByName("creativeAsset")[0].children[0].children[1].children[1].children[0]`
  );
  await RPA.WebBrowser.sendKeys(AssetSearchInput, [`${Datas[0][3]}`]);
  await RPA.sleep(4000);
  const SentakuButton = await RPA.WebBrowser.findElementByXPath(
    `//*[@id="reactroot"]/div/div[5]/div[2]/div[1]/div/form/div/div[2]/div[2]/table/tbody/tr/td[1]/div`
  );
  await SentakuButton.click();
  await RPA.sleep(500);
  // クリエイティブ名を入力する
  const CreativeName: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByName("name")[1].children[1].children[0].children[0]`
  );
  await RPA.WebBrowser.sendKeys(CreativeName, [`${Datas[0][3]}`]);
  await RPA.sleep(300);
  // 訴求を入力する
  const SokyuuInput2 = await RPA.WebBrowser.findElementsByXPath(
    '/html/body/div/div/div[5]/div[2]/div[1]/div/form/div/div[5]/div[2]/div[1]/div/div/div/div/span[1]/div[2]/input'
  );
  await RPA.WebBrowser.sendKeys(SokyuuInput2[0], [Datas[0][6]]);

  const SokyuuFlag = [];
  SokyuuFlag[0] = true;
  const firstSokyuuValue = await RPA.WebBrowser.findElementsByClassName(
    'Select-option'
  );
  const SokyuuValueText = await Promise.all(
    firstSokyuuValue.map(async (elm) => await elm.getText())
  );
  RPA.Logger.info('訴求一覧 → ' + SokyuuValueText);
  for (let i in SokyuuValueText) {
    if (Datas[0][6] == SokyuuValueText[i]) {
      RPA.Logger.info('一致しました' + SokyuuValueText[i]);
      const SokyuuSelectValue = await RPA.WebBrowser.findElementByXPath(
        `/html/body/div/div/div[5]/div[2]/div[1]/div/form/div/div[5]/div[2]/div[1]/div/div/div/div[2]/div/div[${
          Number(i) + Number(1)
        }]`
      );
      await RPA.sleep(100);
      await RPA.WebBrowser.mouseClick(SokyuuSelectValue);
      SokyuuFlag[0] = false;
      await RPA.sleep(500);
    }
  }
  // 訴求が無ければ新規作成を行う
  if (SokyuuFlag[0] == true) {
    const SokyuuCreateButton = await RPA.WebBrowser.findElementByXPath(
      '/html/body/div/div/div[5]/div[2]/div[1]/div/form/div/div[5]/div[2]/div[2]/div/div[1]'
    );
    await RPA.WebBrowser.mouseClick(SokyuuCreateButton);
    await RPA.sleep(200);
    const NewSokyuuInput = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({
        xpath:
          '/html/body/div/div/div[6]/div[2]/div[1]/div/form/div/div[2]/div[2]/div/input',
      }),
      5000
    );
    await RPA.WebBrowser.sendKeys(NewSokyuuInput, [Datas[0][6]]);
    await RPA.sleep(100);
    // 訴求新規作成時にSlackへ通知する
    await RPA.Slack.chat.postMessage({
      channel: SlackChannel,
      token: SlackToken,
      text: `訴求新規作成しました → ${Datas[0][6]}`,
      icon_emoji: ':snowman:',
      username: 'p1',
    });
    const SokyuuOKButton = await RPA.WebBrowser.findElementByXPath(
      '/html/body/div/div/div[6]/div[2]/footer/div[2]'
    );
    await RPA.WebBrowser.mouseClick(SokyuuOKButton);
    await RPA.sleep(4000);
  }
  // 属性が何個あるのか判定して個数によって処理を変える
  const ZokuseiSplitData = await Datas[0][7].split(',');
  if (ZokuseiSplitData.length == 1 && ZokuseiSplitData[0] != '') {
    RPA.Logger.info(`属性の数:1`);
    await ZokuseiInput_function(ZokuseiSplitData[0], true);
  }
  if (ZokuseiSplitData.length == 2) {
    RPA.Logger.info(`属性の数:2`);
    await ZokuseiInput_function(ZokuseiSplitData[0], false);
    await ZokuseiInput_function(ZokuseiSplitData[1], true);
  }
  if (ZokuseiSplitData.length == 3) {
    RPA.Logger.info(`属性の数:3`);
    await ZokuseiInput_function(ZokuseiSplitData[0], false);
    await ZokuseiInput_function(ZokuseiSplitData[1], false);
    await ZokuseiInput_function(ZokuseiSplitData[2], true);
  }
  // 属性無しの場合は、すぐ登録する
  if (ZokuseiSplitData.length == 1 && ZokuseiSplitData[0] == '') {
    RPA.Logger.info('属性 0 なのでスキップします');
    const CreativeOKButto = await RPA.WebBrowser.findElementByXPath(
      '/html/body/div/div/div[5]/div[2]/footer/div[2]'
    );
    await RPA.WebBrowser.mouseClick(CreativeOKButto);
    await RPA.sleep(5000);
  }
}

// 属性を入力する関数
async function ZokuseiInput_function(Data, OKbutton) {
  RPA.Logger.info(`属性 :${Data} 設定します`);
  const CreativeOKButto = await RPA.WebBrowser.findElementByXPath(
    '/html/body/div/div/div[5]/div[2]/footer/div[2]'
  );
  const ZokuseiInput: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByClassName('Select-input')[8].children[0]`
  );
  const ZokuseiNewCreateButton = await RPA.WebBrowser.findElementByXPath(
    '/html/body/div/div/div[5]/div[2]/div[1]/div/form/div/div[6]/div[2]/div[2]/div/div[1]'
  );
  await RPA.sleep(100);

  await RPA.WebBrowser.sendKeys(ZokuseiInput, [Data]);
  await RPA.sleep(300);
  try {
    const list1 = await RPA.WebBrowser.findElementByClassName(
      'Select-menu-outer'
    );
    const listtext1 = await list1.getText();
    // 属性がすでに登録されている場合はそのままOKボタンをおす
    if (listtext1 != 'No results found') {
      RPA.Logger.info('属性既存の物を使用します');
      await RPA.WebBrowser.sendKeys(ZokuseiInput, [Key.ENTER]);
      await RPA.sleep(200);
      if (OKbutton == true) {
        await RPA.WebBrowser.mouseClick(CreativeOKButto);
        await RPA.sleep(5000);
      }
      return;
    }
    // 入力した属性がない場合は、新規作成する
    if (listtext1 == 'No results found') {
      await RPA.WebBrowser.mouseClick(ZokuseiNewCreateButton);
      RPA.Logger.info('属性新規作成します');
      await RPA.sleep(2000);
      const ZokuseiNewInput = await RPA.WebBrowser.findElementByXPath(
        '/html/body/div/div/div[5]/div[2]/div[1]/div/form/div/div[2]/div/input'
      );
      await RPA.WebBrowser.sendKeys(ZokuseiNewInput, [Data]);
      const ZokuseiNewOKButton = await RPA.WebBrowser.findElementByXPath(
        '/html/body/div/div/div[5]/div[2]/footer/div[2]'
      );
      await RPA.WebBrowser.mouseClick(ZokuseiNewOKButton);
      await RPA.sleep(4000);
      if (OKbutton == true) {
        await RPA.WebBrowser.mouseClick(CreativeOKButto);
        await RPA.sleep(5000);
      }
      return;
    }
  } catch {}
}

// クリエイティブIDを取得して　ID円滑シートに貼り付ける関数
async function CreativeIDGet(IDLIST, WorkData, Row) {
  await RPA.sleep(2000);
  await RPA.WebBrowser.refresh();
  await RPA.sleep(1000);
  const firstdata = await RPA.WebBrowser.findElementsByClassName(
    'Table__line___voLKf'
  );
  for (let i in firstdata) {
    const firsttext = await RPA.WebBrowser.findElementByXPath(
      `/html/body/div/div/div[2]/div[3]/div/table/tbody/tr[${
        Number(i) + 1
      }]/td[3]`
    );
    const Text = await firsttext.getText();
    if (Text == WorkData[0][3]) {
      const firstID = await RPA.WebBrowser.findElementByXPath(
        `/html/body/div/div/div[2]/div[3]/div/table/tbody/tr[${
          Number(i) + 1
        }]/td[1]`
      );
      IDLIST[0][1] = await firstID.getText();
      break;
    }
  }
}

// 新しい手法でアセットトランスコードする関数
async function NewAsset_Upload_function(FilePathData) {
  await RPA.sleep(1000);
  const AssetCreateButton = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      className: 'Button__button___2uDT- Button__create___1F14z',
    }),
    5000
  );
  const ButtonText = await AssetCreateButton.getText();
  RPA.Logger.info(ButtonText);
  await RPA.WebBrowser.mouseClick(AssetCreateButton);
  await RPA.sleep(1000);
  const UploaderInput = await RPA.WebBrowser.wait(
    RPA.WebBrowser.Until.elementLocated({
      className: 'Uploader__file___1mcH4',
    }),
    5000
  );
  await RPA.WebBrowser.sendKeys(UploaderInput, [
    `${DownloadFolder}${FilePathData[0]}`,
  ]);
  await RPA.sleep(1000);
  const TransCodeStartButton = await RPA.WebBrowser.findElementByXPath(
    `//*[@id="reactroot"]/div/div[5]/div[2]/div[1]/div/div[2]/div[1]/div[3]`
  );
  await RPA.WebBrowser.mouseClick(TransCodeStartButton);
  await RPA.sleep(1000);
  try {
    for (let i = 0; i < 300; i++) {
      const TransCode = await RPA.WebBrowser.findElementByXPath(
        `//*[@id="reactroot"]/div/div[5]/div[2]/div[1]/div/div[2]/div[1]/div[3]/div/p[1]`
      ); //トランスコード中...の文字
      const TransCodeText = await TransCode.getText();
      RPA.Logger.info(TransCodeText);
      await RPA.sleep(3000);
      if (TransCodeText != 'トランスコード中...') {
        RPA.Logger.info('トランスコード完了');
        break; //トランスコード完了
      }
    }
  } catch {
    RPA.Logger.info('トランスコード完了');
  }
}

// 新しい手法でアセット作成を行う関数
async function NewAsset_Create_function(IDLIST, Datas, Row) {
  await RPA.sleep(2000);
  const AssetNameInput: WebElement = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByName("name")[1].children[1].children[0].children[0]`
  );
  // アセット名が65文字以上なら　円滑シートにエラーを記載して、次の行にスキップする
  if (Datas[0][3].length > 65) {
    const ErrorText = [['エラー', 'アセット名65文字以上']];
    await RPA.Google.Spreadsheet.setValues({
      spreadsheetId: `${SSID}`,
      range: `${SSName1}!K${Row[0]}:L${Row[0]}`,
      values: ErrorText,
    });
    RPA.Logger.info(`${Row[0]} 行目のステータスをエラーに変更しました`);
    await Start();
  }
  if (Datas[0][3].length < 64) {
    RPA.Logger.info('アセット名 65文字以内...OK');
    await AssetNameInput.clear();
    await RPA.sleep(100);
    await RPA.WebBrowser.sendKeys(AssetNameInput, [`${Datas[0][3]}`]);
  }

  // 期間ありか無期限かを判定して、処理を変える
  if (String(Datas[0][5]).length == 0) {
    RPA.Logger.info('無期限にチェックを入れます');
    await RPA.WebBrowser.driver.executeScript(
      `document.getElementsByClassName("Checkboxes__item___2sroS")[0].children[0].click()`
    );
    await RPA.sleep(1000);
  }
  if (String(Datas[0][5]).length > 1) {
    // 有効期間　のボタンをクリックする
    await RPA.WebBrowser.driver.executeScript(
      `document.getElementsByName("dateRange")[0].children[1].children[0].children[0].children[0].children[0].click()`
    );
    // 有効期間の開始時間を記入する
    const TimeRange_start = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({
        name: 'daterangepicker_start',
      }),
      5000
    );
    await RPA.WebBrowser.mouseClick(TimeRange_start);
    await TimeRange_start.clear();
    await RPA.sleep(100);
    await RPA.WebBrowser.sendKeys(TimeRange_start, [Datas[0][4]]);
    await RPA.sleep(300);
    // 有効期間の終了時間を記入する
    const TimeRange_end = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({
        name: 'daterangepicker_end',
      }),
      5000
    );
    await RPA.WebBrowser.mouseClick(TimeRange_end);
    await TimeRange_end.clear();
    await RPA.sleep(100);
    await RPA.WebBrowser.sendKeys(TimeRange_end, [Datas[0][5]]);
    await RPA.sleep(300);
    // 有効期間のOKボタンをおす
    const OKButton: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByClassName("applyBtn btn btn-sm btn-success")[0]`
    );
    await OKButton.click();
    await RPA.sleep(700);
  }
  // 尺が記載されているかどうか判定する
  let SyakuText: string = await RPA.WebBrowser.driver.executeScript(
    `return document.getElementsByName("duration")[1].value`
  );
  RPA.Logger.info(SyakuText);
  const replace1 = await SyakuText.replace('DURATION_', '').replace('S', '');
  RPA.Logger.info(replace1);
  if (replace1 == Datas[0][1]) {
    RPA.Logger.info('尺秒数一致しました');
  }
  if (replace1 != Datas[0][1]) {
    RPA.Logger.info('尺秒数一致しません.エラー処理にてスキップします');
    const ErrorText = [
      ['エラー', '記載されている尺と実際の尺に相違があります。'],
    ];
    await RPA.Google.Spreadsheet.setValues({
      spreadsheetId: `${SSID}`,
      range: `${SSName1}!K${Row[0]}:L${Row[0]}`,
      values: ErrorText,
    });
    await Start();
  }
  if (String(SyakuText).length < 1) {
    RPA.Logger.info('尺が記載されていません');
    // 尺が自動で記載されていなければロボットが記入する
    const SyakuInput: WebElement = await RPA.WebBrowser.driver.executeScript(
      `return document.getElementsByName("duration")[1]`
    );
    await RPA.WebBrowser.mouseClick(SyakuInput);
    await RPA.sleep(100);
    await RPA.WebBrowser.sendKeys(SyakuInput, [Datas[0][1]]);
    await RPA.WebBrowser.sendKeys(SyakuInput, [Key.ENTER]);
    await RPA.sleep(200);
  }
  // アセット名が同じならエラー表示させて、スキップする
  try {
    await RPA.sleep(300);
    const AsseteNameDouble = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementLocated({
        className: 'Field__error___2jFF6',
      }),
      1000
    );
    const AssetNameDoubleText = await AsseteNameDouble.getText();
    if (String(AssetNameDoubleText) == '同じアセット名が既に存在しています') {
      RPA.Logger.info(AssetNameDoubleText + ' 作業スキップします');
      const ErrorText = [['エラー', '同じアセット名が存在しています']];
      await RPA.Google.Spreadsheet.setValues({
        spreadsheetId: `${SSID}`,
        range: `${SSName1}!K${Row[0]}:L${Row[0]}`,
        values: ErrorText,
      });
      await Start();
    }
  } catch {
    RPA.Logger.info('アセット名が唯一なので次の処理に進みます');
  }

  const OKButton = await RPA.WebBrowser.findElementsByClassName(
    `Button__button___2uDT- Button__ok___1qMMy Modal__buttonOkay___2ewPw`
  );
  await RPA.WebBrowser.mouseClick(OKButton[0]); // 登録のボタンをおす
  await RPA.sleep(10000);
}

// アセットIDを取得する関数
async function Get_AssetID_function(IDLIST, Datas, Row) {
  try {
    // アセット一覧ページに遷移
    await RPA.WebBrowser.get(process.env.AAAMS_Account_AssetURL_ALL);
    await RPA.sleep(2000);
    await RPA.WebBrowser.refresh();
    await RPA.sleep(2000);
    // アセットのデータ一覧取得
    const AssetList = await RPA.WebBrowser.wait(
      RPA.WebBrowser.Until.elementsLocated({
        className: 'Table__line___voLKf',
      }),
      7000
    );
    for (let i in AssetList) {
      let AssetText: any = await RPA.WebBrowser.findElementByXPath(
        `//*[@id="reactroot"]/div/div[2]/div[3]/div/table/tbody/tr[${
          Number(i) + 1
        }]/td[3]`
      );
      let AssetID: any = await RPA.WebBrowser.findElementByXPath(
        `//*[@id="reactroot"]/div/div[2]/div[3]/div/table/tbody/tr[${
          Number(i) + 1
        }]/td[1]`
      );
      AssetText = await AssetText.getText();
      AssetID = await AssetID.getText();
      RPA.Logger.info(AssetID, AssetText);
      if (AssetText == Datas[0][3]) {
        RPA.Logger.info('アセット名一致しました.アセットID取得します');
        RPA.Logger.info('アセットID:', AssetID);
        IDLIST[0][0] = AssetID;
        break;
      }
    }
  } catch {
    await RPA.WebBrowser.takeScreenshot();
    RPA.Logger.info('エラー出現しました.RPA終了させます');
    await RPA.sleep(2000);
    await process.exit();
  }
}

async function Spreadsheet_setValues_function(IDLIST, WorkData, Row) {
  // ID円滑シートに　クリエイティブIDを貼り付ける
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName1}!I${Row[0]}:J${Row[0]}`,
    values: IDLIST,
  });
  // 作業終了を記載する
  await RPA.Google.Spreadsheet.setValues({
    spreadsheetId: `${SSID}`,
    range: `${SSName1}!K${Row[0]}:K${Row[0]}`,
    values: [['完了']],
  });
}
