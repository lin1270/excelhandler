
var cmds = require('./cmd/index')

var main = {
    exec (argv) {
        cmds.forEach((cmd)=>{
            cmd.run(argv)
        })
    }
}

module.exports = main