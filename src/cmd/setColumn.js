var Excel = require('node-xlsx')
var utils = require('../utils')
var fs = require('fs')

var cmdName = 'setColumn'

var setColumn = {
    // setColumn excelFileName sheetNameOrIndexFrom0 beginRowFrom0 columnIndexFrom0 txt
    // setColumn a.xls 0 4 1 hehe
    run(argv){
        if (argv[0] !== cmdName) return
        if (argv.length !== 6) {
            console.log('参数不正确。')
            console.log('setColumn excelFileName sheetNameOrIndexFrom0 beginRowFrom0 columnIndexFrom0 txt')
            return
        }
        
        const fileName = argv[1]
        var sheets = Excel.parse(fileName)
        if (!sheets) {
            console.log('打开excel失败：' + fileName)
            return
        }

        let sheetName = argv[2].trim()
        if (utils.isNumber(sheetName)) {
            sheetName = Number(sheetName)
        }

        let sheet = null
        if (typeof sheetName === 'number') {
            if (sheetName >= sheets.length) {
                console.log('sheet name 不正确')
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
            return
        }
        
        
        const beginRowIndex = Number(argv[3])
        const column = Number(argv[4])
        const txt = argv[5]

        sheet.data.forEach((row, index)=>{
            if (index >= beginRowIndex) {
                row[column] = txt
            }
        })
        
        const buffer = Excel.build(sheets)
        // 写入文件
        fs.writeFile(fileName, buffer, function(err) {
            if (err) {
                console.log("保存失败: " + fileName + ' err:' + err);
                return;
            }

            console.log("保存成功：" + fileName);
        });
    }
}


module.exports = setColumn