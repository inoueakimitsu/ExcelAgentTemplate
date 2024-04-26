# ExcelAgentTemplate

## 概要

ExcelAgentTemplate は Excel の関数から LLM を使用したエージェントを実行するための、Excel のアドインと Python のスクリプトのサンプル コードです。

この Excel のアドインを使うと、Excel 上で、プロンプトを文字列として入力し、 LLM を用いたエージェントに指示を与え、得られた結果を文字列として受け取れる関数が使えるようになります。例えば `=RunAgent("企業情報のリサーチャーとして振舞ってください。" A1 & "の所在地を調べてください。")` と入力すると、`A1` セルに記載した企業の所在地をウェブ検索を行い調査し、その結果を出力します。企業名のリストがあれば、すぐに所在地のリストを生成することができます。

### ExcelAgentTemplate の位置づけ

Excel の関数としてエージェントを実行するインターフェースは、チャットでエージェントを実行するインターフェースと、プログラム  コードやフロー ベースのノーコード ツールからエージェントを実行するインターフェースの中間的な用途に適しています。チャット インターフェースでは難しい同一のチャットの繰り返し利用を実現しながら、プログラム コードやフロー ベースのノーコード ツールでは難しい気軽な試行錯誤と、データのアクセス、そして手作業での修正を実現します。

### コードが書ける方向けの解説

同梱の Excel のアドインが行っていることはシンプルで、JSON 形式の入力文をサーバーに送信し、結果をパースして表示しているだけです。同梱の Python のサンプル スクリプトには、LangChain と FastAPI を組み合わせたエージェントの参考実装が含まれます。コードを書ける人は CrewAI、AutoGen、LangChain、LangGraph などを使うことで、ノーコードでも LangFlow、Flowise、Dify、LangServe などを使用することで、用途に応じたエージェントを作成することができます。

## サンプル コードの使用方法

ここでは、サンプル コードを変更せず使う場合の流れを説明します。

### Python

ここでは Windows で uv を使って環境構築する方法を記載します。

- `Python 3.10` 以上をインストールします。
- `uv` をインストールします。
- `uv venv .venv` を実行し仮想環境を作成します。
- `.venv\Scripts\activate` を実行します。
- `uv pip install -r requirements.txt` を実行します。
- `.env.example` を `.env` に別名でコピーし、`.env` をエディターで開き `OPENAI_API_KEY=*****` を自身の OpenAI API Key で書き換えます。 `TAVILY_API_KEY=*****` も自身の Tavily API Key で書き換えます。

インストール後の起動方法は以下の通りです。

- `.venv\Scripts\activate` を実行します。
- `python langchain_fastapi.py` を実行します。
- (Excel の作業が完了したら、Ctrl+C で終了します)

### Excel

- TODO: ビルド済みのバイナリを Relase に追加します。現在はこのリポジトリに含まれていませんので Visual Studio で RunAgentClient.sln を開いてビルドしてください。
- `RunAgentClient/bin/Debug/RunAgentClient-AddIn64.xll` をダブルクリックします。
- "使用できるデジタル署名がありません" という通知画面が表示されます。"このアドインをこのセッションに限り有効にする(E)" をクリックします。このセッションでのみ `RunAgent` 関数が使用できます。
- そのセッションで空白のブックを新規作成してみましょう。
- 任意のセルに、例えば `=RunAgent("株式会社 ??? の従業員数を調べてください。")` と入力し、Enter キーを押下します。??? は適当に置き換えてください。
- `#N/A`と表示されます。Python の画面上では、処理中のログが表示されます。処理が完了すると、セルの内容が実際の出力に置き換わります。

## エージェントを追加するには

- エージェントを実装した API のエンド ポイントを公開します。この際、同一のメッセージが入力された場合のキャッシュ機構を使用すると良いです。
	- 参考 https://python.langchain.com/docs/modules/model_io/llms/llm_caching/
- Excel のアドインは Excel-DNA を使用して作成します。LLM の出力には時間がかかるので、サンプル コードを参考に非同期な処理を行うようにしてください。
	- 参考 https://excel-dna.net/docs/guides-basic/asynchronous-functions
- ウェブ検索に関しては SerpAPI を直に使うことはせず、Tavily を使っています。Tavily は、素の Google 検索や DuckDuckGo、Bing と比べて LLM との相性が良いと印象です。
