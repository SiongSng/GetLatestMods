const CurseForge = require("mc-curseforge-api");
const config = require(`${process.cwd()}/config.json`)  //config
const xlsx = require('node-xlsx'); //試算表解析模塊
const fs = require("fs");
const options = {'!cols': [{ wch: 8 }, { wch: 25 }, { wch: 75 }, { wch: 60 }, { wch: 8 }, { wch: 25 }, { wch: 25 } ]};

let Wdata = [{
    name: '照模組更新時間排序資訊小工具',
    data: [
        [
            '搜索編號',
            '模組名稱',
            '模組敘述',
            '下載網址',
            '下載數量',
            '更新日期',
            '創建日期'
        ],
    ]
},
]
CurseForge.getMods({sort: 2, pageSize: config.PageSize, gameVersion: config.GameVersion}).then((mods) => {
    console.log("-- 照模組更新時間排序資訊小工具 由 菘菘#8663 製作 --\n正在抓取資料中，請稍後...\n提醒您，請不要過量使用此程式以免導致讀取 CurseForge　API 超速。")

    async function aaa() {
        for (let i = 0; i < mods.length; i++) {
            let data = JSON.parse(JSON.stringify(mods[i]))
            if (Date.parse(data.created) > Date.parse(config.Date.split(">")[0]) && Date.parse(data.created) < Date.parse(config.Date.split(">")[1])) {
                Wdata[0].data[Wdata[0].data.length] = [i + 1, data.name, data.summary, data.url, data.downloads, data.updated, data.created];
            }
        console.log(data.logo.url)
        }
    }

    aaa().then(() => {
        let buffer = xlsx.build(Wdata,options);
        fs.writeFile('opt.xlsx', buffer, function (err) {
            if (err) {
                if (err.toString().startsWith("Error: EBUSY: resource busy or locked")) {
                    err = "由於其他程式已經讀取了 opt.xlsx 此檔案，請先關閉該程式後再次執行。"
                }
                console.log("寫入試算表失敗: " + err);
                return;
            }
            console.log("寫入試算表成功。");
        });
    });
});