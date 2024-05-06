# ExcelAgentTemplate

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## Overview

**ExcelAgentTemplate** is a powerful add-in that combines Microsoft Excel with Python. This tool enables users to leverage the latest LLMs (Large Language Models) through Excel functions and execute automated agents. By simply entering specific prompts in Excel cells, users can easily perform complex queries and data processing tasks utilizing LLMs.

For example, using the function `=RunAgent("Act as a corporate information researcher. Please find the location of the company in cell A1.")`, the add-in automatically investigates the location of the company name entered in cell A1 and directly displays the result in Excel. This eliminates the need for manual data searches and input, allowing for efficient work progress.

### Features and Benefits

ExcelAgentTemplate offers the following features:

- **Intuitive Operation**: LLMs can be directly invoked as Excel functions, requiring no special programming knowledge.
- **Advanced Automation**: Complex data collection and processing are automated, significantly streamlining workflows.
- **Customizable**: Agent behavior can be customized according to specific needs using Python scripts.
- **Real-time Processing**: Asynchronous processing enables fast handling of large amounts of data, reflecting results in real-time.

### Use Cases

|     | A        | B                                                                  | C                                                                                                                              |
| --- | -------- | ------------------------------------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------ |
| 1   | Company Name   | Number of Employees - Research Report                                             | Number of Employees                                                                                                                       |
| 2   | Example  | `=RunAgent("Research the number of employees for the company in cell A2")`           | `=RunAgent("Please extract the number of employees from the text in cell B2 and return only the numeric value.")`                  |

### Applicable Scenarios

ExcelAgentTemplate is particularly well-suited for the following purposes:

- **Automating Checklist Verification**: Ideal for checking and correcting inputs based on numerous checklists.
- **Data Collection and Analysis**: Automates the extraction and analysis of necessary information from large datasets.
- **Report Generation**: Simplifies the process of gathering information from multiple data sources, formatting it, and creating reports.
- **Prototyping LLM Products**: Suitable for prototyping and refining products that utilize language models.
- **Translation**: Well-suited for translation tasks using LLMs, especially when combined with human post-editing.

## Information for Developers

### Environment Setup (Python)

1. Install Python 3.10 or higher
2. Execute the following in the command prompt:
   ```bash
   pip install uv
   uv venv .venv
   .venv\Scripts\activate
   pip install -r requirements.txt
   ```
3. Rename `.env.example` to `.env` and set the required API keys

### Environment Setup (Excel)

1. Install Microsoft Excel (2019 or later)
2. Open RunAgentClient.sln in Visual Studio 2019 or later and build the solution.
3. After the build, RunAgentClient/bin/Debug/RunAgentClient-AddIn64.xll will be generated. Double-clicking this file will temporarily enable the Excel add-in.

### Startup Method

1. Activate the virtual environment (`.\.venv\Scripts\activate`) and then run `python langchain_fastapi.py`
2. Double-click `RunAgentClient/bin/Debug/RunAgentClient-AddIn64.xll`.
3. A notification screen stating "There are no usable digital signatures" will be displayed. Click "Enable this add-in for this session only (E)". The `RunAgent` function will be available only for this session.
4. Create a new blank workbook in that session.
5. In any cell, enter a formula like `=RunAgent("Please research the number of employees for XYZ Corporation.")`, replacing "XYZ Corporation" with an appropriate company name, and press Enter.
6. `#N/A` will be displayed. On the Python screen, logs indicating the processing status will be shown. Once the processing is complete, the cell content will be replaced with the actual output.

## Limitations

- Keep track of your OpenAI API usage.
- This tool does not use a GPU.

## Frequently Asked Questions

Q. What tasks is ExcelAgentTemplate suitable for?
A. ExcelAgentTemplate is well-suited for automating various business tasks such as data collection, analysis, and report generation. It is particularly effective for non-standard tasks and situations that require integration with external data.

Q. How much do the APIs cost?
A. API costs depend on usage. Check each API's website for details. Some offer free plans with usage limits.

Q. Is commercial use allowed?
A. Yes, commercial use is permitted. The project is provided under the MIT License, allowing free usage, modification, and distribution. However, please adhere to the terms of service of each API when using them.

Q. How can I add new agents?
A. Please refer to the following:
- Publish an API endpoint that implements the agent. It is recommended to use a caching mechanism for cases where the same message is input.
	- Reference: https://python.langchain.com/docs/modules/model_io/llms/llm_caching/
- The Excel add-in is created using Excel-DNA. Since LLM outputs can take some time, ensure that you perform asynchronous processing as shown in the sample code.
	- Reference: https://excel-dna.net/docs/guides-basic/asynchronous-functions
- For web searches, Tavily is used instead of directly using SerpAPI. Tavily seems to be more compatible with LLMs compared to raw Google searches, DuckDuckGo, or Bing.

If you have any other questions, feel free to ask on the GitHub Issues or the Discord server.

## Community

Discussions and information exchanges related to the project are conducted on the following Discord server:
- [ExcelAgentTemplate Discord Server](https://discord.gg/yCU6DwTX)

Feel free to join!

## Roadmap

The ExcelAgentTemplate project plans to add the following features:

- [ ] Support for Local LLMs
- [ ] Enhancement of documentation

If you have any requests or ideas, please let us know by creating an Issue.

## License

This project is released under the MIT License. Please refer to the LICENSE file for details.

## Contribution

Contributions to the project are welcome. If you have any bug reports or suggestions for additional features, please create an Issue. Pull Requests are also accepted at any time.
