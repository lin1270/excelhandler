
var utils = require('../utils')
var exceljs = require('exceljs')

var cmdName = 'setColumn'

var setColumn = {
    // setColumn excelFileName sheetNameOrIndexFrom1 beginRowFrom1 columnNameOrIndexFrom1 txt
    // setColumn a.xls 1 4 A hehe
    run(argv){
        if (argv[0] !== cmdName) { 
            return
        }
        if (argv.length !== 6) {
            console.log('参数不正确。')
            console.log('setColumn setColumn excelFileName sheetNameOrIndexFrom1 beginRowFrom1 columnNameOrIndexFrom1 txt')
            return
        }
        
        this.runCore(argv[1], argv[2], argv[3], argv[4], argv[5], null)
    },

    runCore(file, sheetStr, beginRow, columnStr, text, cb) {
        var workbook = new exceljs.Workbook()
        workbook.xlsx.readFile(file).then(()=>{
            let sheetName = sheetStr
            if (typeof sheetStr == 'string' && utils.isNumber(sheetName.trim())) {
                sheetName = Number(sheetName.trim())
            }

            let sheet = workbook.getWorksheet(sheetName)
            if (!sheet) {
                console.log('sheet name 不正确')
                if (cb) cb(false)
                return
            }
            
            
            const beginRowIndex = Number(beginRow)
            let column = columnStr
            if (utils.isNumber(columnStr)) column = Number(columnStr)
            const txt = text

            sheet.eachRow((row, index)=>{
                if (index >= beginRowIndex) {
                    row.getCell(column).value = txt
                }
            })
            
            // 写入文件
            workbook.xlsx.writeFile(file)
            .then(function() {
                console.log("保存成功：" + file);
                if (cb) cb(true)
            })
            .catch((e)=>{
                console.log("保存失败: " + file + ' err:' + err);
                if (cb) cb(false)
            })
        }).catch((e)=>{
            console.log('打开excel失败：' + file, e)
            if (cb) cb(false)
        })
    },
}


module.exports = setColumn