Excel VBA と Python を使って、ollama API に非同期にリクエストを行うサンプルプロジェクトです。

## 前提条件

- Windows OS
- ollama Windows Preview のインストール
- `ollama run gemma:7b` などの実行

## セットアップ

1. リポジトリをクローンまたはダウンロードします。

2. Python 3.x をインストールします。

3. 必要な Python のライブラリをインストールします。　pip install requests pywin32

4. Excel ファイルと同じディレクトリに、`async_request.py` をコピーします。

5. Excel ファイルを開き、VBA エディタを開きます（Alt + F11）。

6. リポジトリ内の VBA コードをコピー＆ペーストします。

- `Sheet1` のコードを、Excel ファイルの `Sheet1` オブジェクトに貼り付けます。

7. VBA エディタを閉じ、Excel に戻ります。

8. A 列のセルに入力すると、B 列に API リクエストの結果が表示されます。

## 使い方

1. ollama を起動し、`ollama run gemma:7b` などのコマンドを実行して、API サーバーを起動します。

2. Excel ファイルを開きます。

3. A 列のセルに、ollama API に送信したいプロンプトを入力します。

4. B 列のセルに、ollama API からの応答が表示されます。
