# プログラミング・レッスン４：　トレリスのテキスト描画

## 手順１

👇　以下の内容の 📄 `./temp/lesson/hello_world.json` ファイルを作ってください。  

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
            "xl_pale.xl_red",
            "xl_deep.xl_red"
        ],
        "bgColor": [
            "xl_deep.xl_red",
            "xl_pale.xl_red"
        ]
    },
    "xl_texts": [
        {
            "left": 3,
            "top": 2,
            "width": 1,
            "height": 1,
            "text": "Hello, world!"
        }
    ]
}
```

👆　`["xl_texts"]` の辺りを説明していきます。  
色は趣味で設定してください。  

そして、以下のコマンドを打鍵してください。  

```shell
py trellis.py compile --file ./temp/lesson/hello_world.json --temp ./temp --output ./temp/lesson/hello_world.xlsx
```
