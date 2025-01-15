# Git Hub を見ている人向けの説明


## 環境設定

以下のコマンドを打鍵してください。  

```shell
pip install openpyxl
```


## 動作確認

動作確認用のエグザンプル・ファイルが入っています。  

```shell
# 以下のコマンドで例を全て実行します
py example.py all
```

* 以下のファイルが自動生成されます
    * 📄 `temp/examples/step1_render_empty.xlsx` - step1 で作られる、枠だけ描かれているワークシート
    * 📄 `temp/examples/step2_pillars.xlsx` - step2 で作られる、枠とカードと端子が描かれているワークシート
    * 以下の３ファイルは同じ内容になるはずです
        * 📄 `temp/examples/step3_line_tapes.xlsx` - 枠とカードと端子とラインテープが描かれているワークシート。step3 で作られます
        * 📄 `temp/examples/step4_auto_shadow.xlsx` - step4 で作られます
        * 📄 `temp/examples/step5_auto_split_pillar.xlsx` - step5 で作られます

以上のファイルは、以下のソースファイルから生成されました。  
下の物ほど、簡単になります。  

* 📄 `examples/data/battle_sequence_of_unfair_cointoss.step1_full_manual.json` - step1, 2, 3 で使われます。必要な情報を（自動生成ではなく）全て手入力したものです
* 📄 `examples/data/battle_sequence_of_unfair_cointoss.step4_auto_shadow.json` - step4 で使われます。影の作成を自動設定にしたものです
* 📄 `examples/data/battle_sequence_of_unfair_cointoss.step5_auto_split_by_pillar.json` - step5 で使われます。柱を跨がるラインテープのセグメント作成を自動設定にしたものです


# 実行

📄 `./temp/lesson/hello_world.json` というファイルを作っておくとします。  

自動生成ファイルを入れておくための、 📁 `./temp` というディレクトリーも作っておいてください。  
ファイル名が被って上書きされたり、削除されたりしても困らないフォルダーとして使います。  

以下のコマンドを打鍵してください。  

```shell
py trellis.py compile --level 2 --file ./temp/lesson/hello_world.json --temp ./temp --output ./temp/lesson/hello_world.xlsx
```
