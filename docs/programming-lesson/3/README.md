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
            "baseColor": "xl_light.xl_green"
        }
    ]
}
```

👆　`"rectangles": [` の辺りのコードを書き足してください。  

そして［プログラミング・レッスン２の手順４］と同様に、以下のコマンドを打鍵してください。  

```shell
py trellis.py ruler --file ./temp/lesson/hello_world.json --output ./temp/lesson/hello_world.xlsx
```
