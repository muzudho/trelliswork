# レッスン１：　座標

## 手順１

![列の幅30pixels](../../img/[20250113-1408]column-width-30-pixels.png)  

👆　Excel のワークシートの全ての列の幅を 30 pixels にします。  


## 手順２

![行の高さ30pixels](../../img/[20250113-1411]row-height30pixels.png)  

👆　Excel のワークシートの全ての行の高さを 30 pixels にします。  


## 手順３

![ワークブック・ファイル](../../img/[20250113-1411]no1file.png)  

👆　ファイル名を `no1.xlsx` などにして保存してください。  


## 手順４

以下のコマンドを打鍵してください。

```shell
py trellis.py init
```

そして、指示に従ってください。以下のファイルが作られます。  
📄 `./temp/lesson/hello_world.json`  

ファイルの内容は、例えば以下のようなものです。  

```json
{
    "canvas": {
        "left": 0,
        "top": 0,
        "width": 100,
        "height": 100
    }
}
```

👆　left と top は 0 にしておいてください