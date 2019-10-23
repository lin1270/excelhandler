var setColumn = require('./setColumn')
var fs = require('fs')

var batch = {
    // batch cfg.json
    /**
     * {
     *      cmd: 'setColumn',
     *      args: [
     *          {
     *              file: 'a.xls",
     *              sheet: '0',
     *              beginRow: '0',
     *              column: '0',
     *              text: 'hehe'
     *          }
     *      ]
     * }
     */
    run(argv){
        if (argv[0] !== 'batch') return
        if (argv.length !== 2) {
            console.log('参数不正确。')
            console.log('batch cfg.json')
            return
        }

        const cfgBuffer = fs.readFileSync(argv[1])
        const cfgInfo = JSON.parse(cfgBuffer)

        if (cfgInfo.cmd == 'setColumn') {
            this.runOne(cfgInfo, 0, this.runOne)

            
        }
    },

    runOne(cfg, i, cb) {
        const arg = cfg.args[i]
        if (!arg) {
            console.log('处理完毕')
            return
        }
        setColumn.runCore(arg.file, arg.sheet, arg.beginRow, arg.column, arg.text, (ok)=>{
            if (ok) {
                if (cb) cb(cfg, i + 1, cb)
            }
        })
    }
}

module.exports = batch