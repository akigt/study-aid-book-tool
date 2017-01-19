## docx2html

特定のフォーマットで書かれたdocxをhtmlに変換します。

* 前提

rubyzipのgemが必要です。

入っていない場合は、

```sh
$ gem install rubyzip
```

を実行するか、Gemfileを作成してください。

* 使い方

```text
study-aid-book-tool/
├ docx2html.rb
├ docx/
｜  ├ 元となるファイル.docx
｜  └ 元となるファイル.docx
└ html/
└ xml/
```

docx/ フォルダにhtmlにしたいdocxを入れます。

```sh
ruby docx2html.rb
```
を実行します。

* 結果

html/ フォルダ内にhtmlが出力されます。
