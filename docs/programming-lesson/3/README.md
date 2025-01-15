# プログラミング・レッスン３：　トレリスの矩形描画

## 手順１

👇　［プログラミング・レッスン２の手順４］で作った 📄 `./temp/lesson/hello_world.json` ファイルの内容について、  

```json
{
    "canvas": {
        "left": 0,
        "top": 0,
        "width": 10,
        "height": 10
    },
    "ruler": {
        "visible": true,
        "fgColor": [
            "xl_deep.xl_red",
            "xl_deep.xl_green",
            "xl_deep.xl_blue"
        ],
        "bgColor": [
            "xl_pale.xl_red",
            "xl_pale.xl_green",
            "xl_pale.xl_blue"
        ]
    },
    "rectangles": [
        {
            "left": 3,
            "top": 2,
            "width": 4,
            "height": 1,
            "bgColor": "xl_light.xl_green"
        }
    ]
}
```

👆　`"rectangles": [` の辺りのコードを書き足してください。  

そして、以下のコマンドを打鍵してください。  

```shell
py trellis.py compile --level 0 --file ./temp/lesson/hello_world.json --temp ./temp --output ./temp/lesson/hello_world.xlsx
```

![矩形描画](../../img/[20250116-0015]rectangle.png)  

👆　［矩形］を描画できました。  


## 手順２

👇　引き続き 📄 `./temp/lesson/hello_world.json` ファイルの内容について、  

```json
{
    "canvas": {
        "left": 0,
        "top": 0,
        "width": 10,
        "height": 10
    },
    "ruler": {
        "visible": true,
        "fgColor": [
            "xl_deep.xl_red",
            "xl_deep.xl_green",
            "xl_deep.xl_blue"
        ],
        "bgColor": [
            "xl_pale.xl_red",
            "xl_pale.xl_green",
            "xl_pale.xl_blue"
        ]
    },
    "rectangles": [
        {
            "left": 3,
            "right": 7,
            "top": 2,
            "bottom": 3,
            "bgColor": "xl_light.xl_blue"
        }
    ]
}
```

👆　`["rectangles"]["width"]` に代えて `["rectangles"]["right"]` を、  
`["rectangles"]["height"]` に代えて `["rectangles"]["bottom"]` を使ってください。  

そして、以下のコマンドを打鍵してください。  

```shell
py trellis.py compile --level 0 --file ./temp/lesson/hello_world.json --temp ./temp --output ./temp/lesson/hello_world.xlsx
```

![右と下を使って矩形描画](../../img/[20250116-0020]right-bottom.png)  

👆　手順１と同じサイズの［矩形］を描画できました。  
