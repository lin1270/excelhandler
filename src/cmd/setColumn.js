var Excel = require('node-xlsx')
var utils = require('../utils')
var fs = require('fs')

var cmdName = 'setColumn'

var setColumn = {
    // setColumn excelFileName sheetNameOrIndexFrom0 beginRowFrom0 columnIndexFrom0 txt
    // setColumn a.xls 0 4 1 hehe
    run(argv){
        if (argv[0] !== cmdName) { 
            return
        }
        if (argv.length !== 6) {
            console.log('参数不正确。')
            console.log('setColumn excelFileName sheetNameOrIndexFrom0 beginRowFrom0 columnIndexFrom0 txt')
            return
        }
        
        const fileName = argv[1]
        this.runCore(argv[1], argv[2], argv[3], argv[4], argv[5], null)
    },

    runCore(file, sheetStr, beginRow, column, text, cb) {
        var sheets = Excel.parse(file)
        if (!sheets) {
            console.log('打开excel失败：' + file)
            if (cb) cb(false)
            return
        }

        let sheetName = sheetStr.trim()
        if (utils.isNumber(sheetName)) {
            sheetName = Number(sheetName)
        }

        let sheet = null
        if (typeof sheetName === 'number') {
            if (sheetName >= sheets.length) {
                console.log('sheet name 不正确')
                if (cb) cb(false)
                return
            }
            sheet = sheets[sheetName]
        } else {
            sheet = sheets.find((si)=>{
                return si.name === sheetName
            })
        }
        if (!sheet) {
            console.log('sheet name 不正确')
            if (cb) cb(false)
            return
        }
        
        
        const beginRowIndex = Number(beginRow)
        column = Number(column)
        const txt = text

        sheet.data.forEach((row, index)=>{
            if (index >= beginRowIndex) {
                row[column] = txt
            }
        })
        
        const buffer = Excel.build(sheets)
        // 写入文件
        fs.writeFile(file, buffer, function(err) {
            if (err) {
                console.log("保存失败: " + file + ' err:' + err);
                if (cb) cb(false)
                return;
            }

            console.log("保存成功：" + file);
            if (cb) cb(true)
        });
    }
}


module.exports = setColumn