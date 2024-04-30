**ExcelAgentTemplate: Unlock the Power of LLM Agent in Excel**

**Overview**

ExcelAgentTemplate is a revolutionary Excel add-in and Python script that lets you harness the power of Large Language Models (LLMs) to automate tasks and generate insights directly within Excel.

With this add-in, you can create custom functions that take a prompt as input, execute an LLM-powered agent, and return the results as a string output. For example, `=RunAgent("Act as a corporate information researcher. Find the location of " & A1 & ".")` will search the web for the location of the company listed in cell `A1` and output the result.

**Example Usage**

|     | A        | B                                                                  | C                                                                                                                              |
| --- | -------- | ------------------------------------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------ |
| 1   | Company  | NumEmployeesResearchResult                                         | NumEmployees                                                                                                                   |
| 2   | Example | =**RunAgent("Report the number of employees of " & A2)** | =**RunAgent(“From the following text, extract the number of employees and output the number only: ” & B2)** |

**Ideal Use Cases**

ExcelAgentTemplate is perfect for non-routine tasks that require flexibility and creativity. It offers an ideal interface for running agents as Excel functions, allowing you to:

* Easily experiment with different prompts and inputs
* Access and manipulate data within Excel
* Make manual modifications and adjustments as needed

**Technical Details**

The included Excel add-in sends a JSON-formatted input string to the server, parses the result, and displays it. The Python sample script provides a reference implementation of an agent combining LangChain and FastAPI. You can create custom agents using various tools and libraries, such as CrewAI, AutoGen, LangChain, LangGraph, and more.

**Getting Started**

### Python

1. Install Python 3.10 or higher and `uv`.
2. Create a virtual environment and install the required packages.
3. Copy the `.env.example` file and replace the API keys with your own.
4. Run the Python script to start the server.

### Excel

1. Build the `RunAgentClient` solution in Visual Studio.
2. Double-click the built binary to enable the add-in.
3. Create a new blank workbook and enter a formula like `=RunAgent("Find the number of employees of ??? Co., Ltd.")`.
4. Press Enter to execute the agent and display the output.

**Tips and Best Practices**

* Publish your API endpoint and consider using caching for repeated inputs.
* Use asynchronous processing to handle LLM outputs efficiently.
* For web searches, try using Tavily for better results with LLMs.
