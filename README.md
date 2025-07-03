# LlmDocxTrans

LlmDocxTrans 工具 - 使用大语言模型翻译 DOCX 文件

用法:
  LlmDocxTrans.exe llm=[model] apikey=[your_api_key] input=[file_path] [其他选项]

必需参数:
  llm=[model]       指定要使用的大语言模型，支持的值:
                      deepseek - DeepSeek 模型
                      doubao   - 豆包模型
                      chatgpt  - ChatGPT 模型
  apikey=[key]      对应模型的 API 密钥
  input=[path]      输入 DOCX 文件的路径

示例:
  LlmDocxTrans.exe llm=doubao apikey="your_doubao_key" input="c:\docs\report.docx"
