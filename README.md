# excelhandler

目前仅支持设置某一列的值。

安装：
```
npm install excelhandler -g
```

使用：
```
excelhandler 
```

直接打命令，会弹出使用说明
```
setColumn excelFileName sheetNameOrIndexFrom1 beginRowFrom1 columnNameOrIndexFrom1 txt
    e.g.    setColumn a.xls 1 4 A hehe
```


批量修改：
```
batch cfg.json
```

cfg.json:
```
{
	"cmd":"setColumn",
	"args":[
		{
			"file":"a.xlsx",
			"sheet":"工作表1",
			"beginRow":"4",
			"column":"G",
			"text":"-hehe-"
		},
		{
			"file":"a.xlsx",
			"sheet": 1,
			"beginRow":2,
			"column":"AA",
			"text":"?hehe?"
		}
	]
}
```