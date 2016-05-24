# arrange_excel.exe.rb

## ファイル種別

実行ファイル

## 必須

- Windows OS
- Excelアプリケーション
- arrange_excel.rb(同じディレクトリに配置する)
- Rubyが実行できる環境(Rubyコマンドプロンプト、Rubyを入れたCygwin等)

## コマンド例

```
ruby arrange_excel.exe.rb
```
->実行フォルダ内のファイルをすべて処理する

```
ruby arrange_excel.exe.rb file1.xlsx file2.xls dir/
```
->指定ファイル、または指定ディレクトリ内のファイルすべてを処理する

```
ruby arrange_excel.exe.rb -r dir/
```
->指定したディレクトリをサブディレクトリも含め、再帰的に処理する

```
ruby arrange_excel.exe.rb -v
```
->実行フォルダ内のファイルをすべて処理し、処理経過を表示する

