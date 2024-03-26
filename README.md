# excel-vba-python-async-api

ollama windows previewをインストールし、ollama run gemma:7bなどを実行したのを前提とします。


セットアップ

1.リポジトリをクローンまたはダウンロードします。
2.Python 3.x をインストールします。
3.必要な Python のライブラリをインストールします。
```bash
pip install requests pywin32
```
4.Excel ファイルと同じディレクトリに、async_request.py をコピーします。
5.Excel ファイルを開き、VBA エディタを開きます。
6.リポジトリ内の Sheet1 コードをコピー＆ペーストします。
7.VBA エディタを閉じ、Excel に戻ります。
8.A 列のセルに入力すると、B 列に API リクエストの結果が表示されます。
