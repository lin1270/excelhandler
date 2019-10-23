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
setColumn excelFileName sheetNameOrIndexFrom0 beginRowFrom0 columnIndexFrom0 txt
    e.g.    setColumn a.xls 0 4 1 hehe
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
			"sheet":"0",
			"beginRow":"2",
			"column":"3",
			"text":"hehe"
		},
		{
			"file":"a.xlsx",
			"sheet":"0",
			"beginRow":"2",
			"column":"4",
			"text":"hehe"
		}
	]
}
```