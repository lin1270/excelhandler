
var help = {
    // setColumn excelFileName sheetNameOrIndexFrom0 beginRowFrom0 columnIndexFrom0 txt
    // setColumn a.xls 0 4 1 hehe
    run(argv){
        const cmdName = argv[0]
        if (!cmdName || cmdName === '-h' || cmdName === '-help') {
            console.log('excelhandler setColumn excelFileName sheetNameOrIndexFrom1 beginRowFrom1 columnNameOrIndexFrom1 txt')
            console.log('    e.g.    excelhandler setColumn a.xls 1 4 A hehe')
        }
    }
}

module.exports = help