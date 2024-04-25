"""LangChain でウェブ検索可能な Agent を実行する Web API を提供します。"""
from fastapi import FastAPI
import uvicorn
import argparse
from langchain_openai.chat_models import ChatOpenAI
from langchain_core.globals import set_debug
from langchain_community.tools.tavily_search import TavilySearchResults

from langchain.agents import create_openai_tools_agent, AgentExecutor
from langchain import hub
from dotenv import load_dotenv
from pydantic import BaseModel, Field
from joblib import Memory

# デフォルトのポート番号を設定します。
DEFAULT_PORT: int = 8889

# デフォルトのホストを設定します。
DEFAULT_HOST: str = "0.0.0.0"

# デフォルトのログレベルを設定します。
DEFAULT_LOG_LEVEL: str = "debug"

# デフォルトのキャッシュディレクトリを設定します。
DEFAULT_CACHE_DIR: str = "./langchain_whole_chain_cache"

# 実装メモ:
# LangChain の Caching 機能は LLM の入出力はキャッシュできても、
# Chain 全体はキャッシュできません。
# このため joblib の Memory を使用してキャッシュを行います。
# なお、同じ入力でも実行毎に出力が変わることが良い場合は
# キャッシュを無効にしたり、キャッシュの有効期限を設定した方がよいでしょう。
# ユースケースに応じて実装を選択してください。
whole_chain_cache_memory = Memory(DEFAULT_CACHE_DIR)

# .env ファイルが存在する場合は読み込み環境変数を設定します。
load_dotenv()

# デバッグ表示を有効にします。
set_debug(True)

# FastAPI のインスタンスを作成します。
app = FastAPI()


# FastAPI の API の仕様を決めるための Pydantic モデルを定義します。
class ChatInput(BaseModel):
    """
    FastAPI の POST リクエストのボディとして使用される Pydantic モデルです。
    ユーザーからのメッセージを保持します。

    Attributes:
        message (str): ユーザーからのメッセージです。空文字は許可されません。
        model (str): 使用するモデルの名前です。デフォルトは "gpt-4-turbo-preview" です。
    """
    message: str = Field(
        description="ユーザーからのメッセージです。空文字は許可されません。")
    model: str = Field(
        "gpt-4-turbo-preview",
        description="使用するモデルの名前です。デフォルトは 'gpt-4-turbo-preview' です。")


@whole_chain_cache_memory.cache
async def chat_internal(chat_input: ChatInput):
    """
    ユーザーからのメッセージを文字列で受け取り、AI エージェントが応答を生成します。

    Parameters:
        chat_input (ChatInput): ユーザーからのメッセージを含む ChatInput オブジェクトです。

    Returns:
        str: 応答メッセージ。メッセージが空の場合は空文字を返します。
    """
    # メッセージが空の場合は空文字を返します。
    if chat_input.message == "":
        return ""

    # ウェブ検索に使うツールを設定します。
    web_search_tools = [
        TavilySearchResults(
            description=(
                "A search engine optimized for comprehensive, accurate, and trusted results. "
                "Useful for when you need to answer questions about current events. "
                "Input should be a search query. "
                "If you don't get good search results, "
                "please change the keywords and search again."
            ),
            max_results=10,
            verbose=True),
    ]

    # エージェントが使用するツールを設定します。
    # 実装メモ: 他のツールを使用する場合は、ここに追加してください。
    tools = web_search_tools

    # ウェブ検索用のエージェントを作成します。
    agent = create_openai_tools_agent(
        llm=ChatOpenAI(model=chat_input.model),
        tools=tools,
        prompt=hub.pull("hwchase17/openai-tools-agent")
    )

    # エージェントを Runnable に変換します。
    chain = AgentExecutor(
        agent=agent, tools=tools,
        handle_parsing_errors=True,
        max_iterations=30, verbose=True)

    # エージェントにメッセージを送信して返答を取得します。
    result = await chain.ainvoke({"input": chat_input.message})

    return result["output"]


@app.post("/chat")
async def chat(chat_input: ChatInput):
    """Agent にメッセージを送信して返答を取得するエンドポイントです。

    キャッシュが存在する場合はキャッシュを使用して返答を取得します。
    """
    result = await chat_internal(chat_input)
    return result


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="LangChain Agent on FastAPI")
    parser.add_argument('--host', default=DEFAULT_HOST, help='Host to run the app on')
    parser.add_argument('--port', type=int, default=DEFAULT_PORT, help='Port to run the app on')
    parser.add_argument('--log_level', default=DEFAULT_LOG_LEVEL, help='Logging level')
    args = parser.parse_args()
    
    # 公開するホストとポートを指定して FastAPI を起動します。
    uvicorn.run(app, host=args.host, port=args.port, log_level=args.log_level)
