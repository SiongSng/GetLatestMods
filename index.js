console.log("-- CurseForge模組列表器 由 菘菘#8663 製作 --\n正在抓取資料中，請稍後...\n提醒您，請不要過量使用此程式以免導致API使用過量。")

const CurseForge = require("mc-curseforge-api");
const request = require("request");
const Excel = require('excel4node');
const fs = require("fs");
const path = require("path");
const config = require(`${process.cwd()}/config.json`)  //config
const options = {
};
const translate = require('@vitalets/google-translate-api');
let TranslationProgress = 0;
let num = 0;

function delDir(path) { //資料夾/檔案迴圈刪除 程式碼來自:https://www.itread01.com/content/1541387043.html
    let files = [];
    if (fs.existsSync(path)) {
        files = fs.readdirSync(path);
        files.forEach((file) => {
            let curPath = path + "/" + file;
            if (fs.statSync(curPath).isDirectory()) {
                delDir(curPath); //遞迴刪除資料夾
            } else {
                fs.unlinkSync(curPath); //刪除檔案
            }
        });
        fs.rmdirSync(path);
    }
}


let wb = new Excel.Workbook();
let ws = wb.addWorksheet('模組資料表格', options);

let style = wb.createStyle({ //試算表格式
    font: {
        color: '#000000',
        size: 14,
    },
});
ws.cell(1, 1).string("模組圖示").style(style)
ws.cell(1, 2).string("搜索編號").style(style)
ws.cell(1, 3).string("模組名稱").style(style)
ws.cell(1, 4).string("模組敘述").style(style)
ws.cell(1, 5).string("模組敘述(機器翻譯)").style(style)
ws.cell(1, 6).string("下載網址").style(style)
ws.cell(1, 7).string("下載數量").style(style)
ws.cell(1, 8).string("更新日期").style(style)
ws.cell(1, 9).string("創建日期").style(style)

let dirPath = path.join(__dirname, "icon"); //暫存模組圖示的位置
delDir(dirPath)
if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath);
}

CurseForge.getMods({sort: 2, pageSize: config.PageSize, gameVersion: config.GameVersion}).then((mods) => {

    async function aaa() {
        console.log("正在翻譯模組敘述中，請稍後...")
        for (let i = 0; i < mods.length; i++) {
            let data = JSON.parse(JSON.stringify(mods[i]));
            if (Date.parse(data.created) > Date.parse(config.Date.split(">")[0]) && Date.parse(data.created) < Date.parse(config.Date.split(">")[1])) {
                num++
                let Translated;
                await translate(data.summary, {to: 'zh-TW'}).then(res => {
                    console.log(`翻譯進度: ${TranslationProgress / num * 100}%`)
                    TranslationProgress++
                    Translated = res.text;
                }).catch(err => {
                    console.error(err);
                });
                for (let k = 0; k < 1; k++) {
                    let stream = fs.createWriteStream(path.join(`./icon/${data.logo.url.toString().substr(43, 65)}`));
                    await request(data.logo.url).pipe(stream).on("close", function (err) {
                         ws.addImage({ //將模組圖片新增到試算表內
                            path: `./icon/${data.logo.url.toString().substr(43, 65)}`,
                            type: 'picture',
                             position: {
                                 type: 'twoCellAnchor',
                                 from: {
                                     col: 1,
                                     colOff: 0,
                                     row: 1,
                                     rowOff: 0,
                                 },
                                 to: {
                                     col: 2,
                                     colOff: 0,
                                     row: 5,
                                     rowOff: 0,
                                 },
                             },
                        });
                    });
                }
                ws.cell(num + 1, 2).number(i + 1).style(style);
                ws.cell(num + 1, 3).string(data.name).style(style);
                ws.cell(num + 1, 4).string(data.summary).style(style);
                ws.cell(num + 1, 5).string(Translated).style(style);
                ws.cell(num + 1, 6).link(data.url).style(style);
                ws.cell(num + 1, 7).number(data.downloads).style(style);
                ws.cell(num + 1, 8).string(data.updated).style(style);
                ws.cell(num + 1, 9).string(data.created).style(style);
            }
        }
    }

    aaa().then(() => {
        console.log(`翻譯進度: 100%`)
        wb.write('opt.xlsx', function (err) {
            if (err) {
                console.error(err);
            } else {
                console.log("成功寫入試算表"); // Prints out an instance of a node.js fs.Stats object
            }
        });
    });
});