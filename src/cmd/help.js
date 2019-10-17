
var help = {
    // setColumn excelFileName sheetNameOrIndexFrom0 beginRowFrom0 columnIndexFrom0 txt
    // setColumn a.xls 0 4 1 hehe
    run(argv){
        const cmdName = argv[0]
        if (!cmdName || cmdName === '-v' || cmdName === '-help') {
            console.log('setColumn excelFileName sheetNameOrIndexFrom0 beginRowFrom0 columnIndexFrom0 txt')
            console.log('    e.g.    setColumn a.xls 0 4 1 hehe')
        }
    }
}

module.exports = help