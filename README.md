# ExcelAgentTemplate

## Overview

ExcelAgentTemplate is a sample code for an Excel add-in and Python script that allows you to run agents using LLMs from Excel functions.

With this Excel add-in, you can use functions in Excel that take a prompt as a string input, give instructions to an agent using LLMs, and receive the results as a string output. For example, if you enter `=RunAgent("Act as a corporate information researcher. Find the location of " & A1 & ".")`, it will search the web for the location of the company listed in cell `A1` and output the result. If you have a list of company names, you can quickly generate a list of locations.

For Example:
|     | A        | B                                                                  | C                                                                                                                              |
| --- | -------- | ------------------------------------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------ |
| 1   | Company  | NumEmployeesResearchResult                                         | NumEmployees                                                                                                                   |
| 2   | Example | =**RunAgent("Report the number of employees of " & A2)** | =**RunAgent(“From the following text, extract the number of employees and output the number only: ” & B2)** |


### Use Cases Suitable for ExcelAgentTemplate

We recommend using ExcelAgentTemplate primarily for non-routine tasks.

The interface for running agents as Excel functions is suitable for intermediate use cases between running agents through a chat interface and running agents from program code or flow-based no-code tools. It achieves the repeated use of the same chain, which is difficult with chat interfaces, while enabling easy trial and error, data access, and manual modifications, which are difficult with program code or flow-based no-code tools.

The fact that the same chain is executed multiple times is an advantage of using the chain abstraction even if you have to operate a time-consuming UI that is not in chat format. However, in non-routine tasks, it is rare to use the same chain the next day. This is why flow-based UIs are not suitable for some tasks. An Excel-based UI is ideal for this use case.

The granularity of function calls is also important. In daily work, you don't want to make calls at the same granular level as LLM utterances. We believe that the final execution result of an agent is the just-right granularity for function calls. This is another point where ExcelAgentTemplate can contribute.

When non-routine tasks become routine, let's replace this with LangChain or similar tools. We are developing a tool to convert from Excel to LangChain.

### Explanation for Those Who Can Write Code

What the included Excel add-in does is simple: it sends a JSON-formatted input string to the server, parses the result, and displays it. The included Python sample script contains a reference implementation of an agent combining LangChain and FastAPI. Those who can write code can create agents for various purposes using CrewAI, AutoGen, LangChain, LangGraph, etc., or using no-code tools such as LangFlow, Flowise, Dify, LangServe, etc.

## How to Use the Sample Code

Here, we explain the process of using the sample code without modification.

### Python

Here, we describe how to set up the environment using uv on Windows.

- Install `Python 3.10` or higher.
- Install `uv`.
- Run `uv venv .venv` to create a virtual environment.
- Run `.venv\Scripts\activate`.
- Run `uv pip install -r requirements.txt`.
- Copy `.env.example` to a new file named `.env`, open `.env` in an editor, and replace `OPENAI_API_KEY=*****` with your own OpenAI API Key. Also replace `TAVILY_API_KEY=*****` with your own Tavily API Key.

After installation, follow these steps to start:

- Run `.venv\Scripts\activate`.
- Run `python langchain_fastapi.py`.
- (When you're done working in Excel, press Ctrl+C to exit)

### Excel

- TODO: Add the built binary to the Release. Currently, it is not included in this repository, so please open RunAgentClient.sln in Visual Studio and build it.
- Double-click `RunAgentClient/bin/Debug/RunAgentClient-AddIn64.xll`.
- A notification screen saying "No valid digital signature is available" will appear. Click "Enable this add-in for this session only (E)". The `RunAgent` function will be available only for this session.
- Try creating a new blank workbook in that session.
- In any cell, enter a formula like `=RunAgent("Find the number of employees of ??? Co., Ltd.")`, replacing ??? with an appropriate company name, and press Enter.
- `#N/A` will be displayed. On the Python screen, logs of the processing will be shown. When the processing is complete, the cell content will be replaced with the actual output.

## Tips

- Publish the endpoint of the API that implements the agent. In doing so, it's a good idea to use a caching mechanism for cases where the same message is input.
	- Reference: https://python.langchain.com/docs/modules/model_io/llms/llm_caching/
- The Excel add-in is created using Excel-DNA. Since LLM outputs take time, make sure to perform asynchronous processing as shown in the sample code.
	- Reference: https://excel-dna.net/docs/guides-basic/asynchronous-functions
- For web searches, we recommend to use Tavily. In my impression, Tavily works better with LLMs compared to raw Google search, DuckDuckGo, or Bing.
