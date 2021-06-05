console.log("-- CurseForge模組列表器 由 菘菘#8663 製作 --\n正在抓取資料中，請稍後...\n提醒您，請不要過量使用此程式以免導致API使用過量。")

const CurseForge = require("mc-curseforge-api");
const request = require("request");
const Excel = require('excel4node');
const fs = require("fs");
const path = require("path");
const config = require(`${process.cwd()}/config.json`)  //config
const translate = require('@vitalets/google-translate-api');
let TranslationProgress = 0;
let num = 0;

if(config.PageSize > 1000){
    return console.log("由於您輸入的抓取數值大於 1000 ，可能會導致API使用過量，因此系統自動停止此操作。")
}

async function Translate(scr) {
    let opt;
    await translate(scr, {to: 'zh-TW'}).then(res => {
        opt = res.text
            .replace("暴民", "生物")
            .replace("暴徒", "生物")
            .replace("Mods", "模組")
            .replace("mods", "模組")
            .replace("MODS", "模組")
            .replace("Mod", "模組")
            .replace("mod", "模組")
            .replace("MOD", "模組")
            .replace("支持", "支援")
            .replace("XP", "經驗值")
            .replace("香草", "原版")
            .replace("老闆", "BOSS");
    }).catch(err => {
        console.error(err);
    });
    return opt;
}

function addImage(url, index) {
    let stream = fs.createWriteStream(path.join(`./icon/${url.toString().substr(43, 65)}`));
    request(url).pipe(stream).on("close", function (err) {
        if (err) throw err;
        ws.row(index + 1).setHeight(30);
        try {
            ws.addImage({
                path: `./icon/${url.toString().substr(43, 65)}`,
                type: 'picture',
                position: {
                    type: 'twoCellAnchor',
                    from: {
                        col: 1,
                        colOff: 0,
                        row: index + 1,
                        rowOff: 0,
                    },
                    to: {
                        col: 2,
                        colOff: 0,
                        row: index + 2,
                        rowOff: 0,
                    },
                },
            });
        } catch (err) {
            ws.cell(index + 1, 1).string("無效").style(style);
        }
    })
}

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
let ws = wb.addWorksheet('模組資料表格');

ws.column(1).setWidth(15 / 3);
ws.column(2).setWidth(30);
ws.column(3).setWidth(30);
ws.column(4).setWidth(75);
ws.column(5).setWidth(60);
ws.column(6).setWidth(80);
ws.column(7).setWidth(10);
ws.column(8).setWidth(35);
ws.column(9).setWidth(35);


let style = wb.createStyle({ //試算表格式
    font: {
        color: '#000000',
        size: 14,
    },
});
ws.cell(1, 2).string("模組名稱").style(style)
ws.cell(1, 3).string("模組名稱(機器翻譯)").style(style)
ws.cell(1, 4).string("模組敘述").style(style)
ws.cell(1, 5).string("模組敘述(機器翻譯)").style(style)
ws.cell(1, 6).string("下載網址").style(style)
ws.cell(1, 7).string("下載數量").style(style)
ws.cell(1, 8).string("更新日期").style(style)
ws.cell(1, 9).string("創建日期").style(style)

delDir("./icon")
if (!fs.existsSync("./icon")) {
    fs.mkdir("./icon", function (err) {
        if (err) throw err;
    });
}
CurseForge.getMods({sort: 2, pageSize: config.PageSize * 10, gameVersion: config.GameVersion}).then((mods) => {

        async function Run() {
            console.log("正在翻譯模組敘述中，請稍後...")
            for (let i = 0; i < mods.length; i++) {
                let data = JSON.parse(JSON.stringify(mods[i]));
                if (Date.parse(data.created) > Date.parse(config.Date.split(">")[0]) && Date.parse(data.created) < Date.parse(config.Date.split(">")[1])) {
                    num++
                    if (num >= config.PageSize) break;
                    await addImage(data.logo.url, num);
                    console.log(`翻譯進度: ${TranslationProgress / num * 100}%`)
                    TranslationProgress++;
                    ws.cell(num + 1, 2).string(data.name).style(style);
                    ws.cell(num + 1, 3).string(await Translate(data.name)).style(style);
                    ws.cell(num + 1, 4).string(data.summary).style(style);
                    ws.cell(num + 1, 5).string(await Translate(data.summary)).style(style);
                    ws.cell(num + 1, 6).link(data.url).style(style);
                    ws.cell(num + 1, 7).number(data.downloads).style(style);
                    ws.cell(num + 1, 8).string(data.updated).style(style);
                    ws.cell(num + 1, 9).string(data.created).style(style);
                }
            }
        }

        Run().then(() => {
            console.log(`翻譯進度: 100%`)
            wb.write('opt.xlsx', function (err) {
                if (err) {
                    console.error(err);
                } else {
                    console.log("成功寫入試算表"); // Prints out an instance of a node.js fs.Stats object
                }
            });
        });
    }
);