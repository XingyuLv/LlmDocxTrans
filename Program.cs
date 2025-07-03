// See https://aka.ms/new-console-template for more information


if (args.Length == 0 || args[0].ToLower() == "--help" || args[0].ToLower() == "-h")
{
    ShowHelp();
    return;
}

try
{
    var options = ParseCommandLineArgs(args);

    // 验证必需参数
    if (!options.ContainsKey("llm") || !options.ContainsKey("apikey") || !options.ContainsKey("input"))
    {
        Console.WriteLine("错误：缺少必需的参数。");
        ShowHelp();
        return;
    }

    var validLlms = new[] { "deepseek", "doubao", "chatgpt" };
    if (!validLlms.Contains(options["llm"].ToLower()))
    {
        Console.WriteLine($"错误：不支持的 LLM 类型。支持的类型：{string.Join(", ", validLlms)}");
        ShowHelp();
        return;
    }

    // 处理文件路径中的引号
    if (options["input"].StartsWith("\"") && options["input"].EndsWith("\""))
        options["input"] = options["input"][1..^1];

    Console.WriteLine($"使用 {options["llm"]} 翻译文件：{options["input"]}");
    LlmDocxTrans.Process(options["input"],options["llm"],options["apikey"]);

}
catch (Exception ex)
{
    Console.WriteLine(ex);
    Console.WriteLine($"处理命令行参数时出错：{ex.Message}");
    ShowHelp();
}


Dictionary<string, string> ParseCommandLineArgs(string[] args)
{
    var options = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

    foreach (var arg in args)
    {
        var parts = arg.Split(new[] { '=' }, 2);
        if (parts.Length == 2)
            options[parts[0].Trim()] = parts[1].Trim();
    }

    return options;
}

void ShowHelp()
{
    Console.WriteLine("LlmDocxTrans 工具 - 使用大语言模型翻译 DOCX 文件");
    Console.WriteLine();
    Console.WriteLine("用法:");
    Console.WriteLine("  LlmDocxTrans.exe llm=[model] apikey=[your_api_key] input=[file_path] [其他选项]");
    Console.WriteLine();
    Console.WriteLine("必需参数:");
    Console.WriteLine("  llm=[model]       指定要使用的大语言模型，支持的值:");
    Console.WriteLine("                      deepseek - DeepSeek 模型");
    Console.WriteLine("                      doubao   - 豆包模型");
    Console.WriteLine("                      chatgpt  - ChatGPT 模型");
    Console.WriteLine("  apikey=[key]      对应模型的 API 密钥");
    Console.WriteLine("  input=[path]      输入 DOCX 文件的路径");
    Console.WriteLine();
    // Console.WriteLine("可选参数:");
    // Console.WriteLine("  source=[lang]     源语言代码 (默认: 自动检测)");
    // Console.WriteLine("  target=[lang]     目标语言代码 (默认: zh-CN)");
    Console.WriteLine();
    Console.WriteLine("示例:");
    Console.WriteLine("  LlmDocxTrans.exe llm=doubao apikey=\"your_doubao_key\" input=\"c:\\docs\\report.docx\"");
    // Console.WriteLine("  LlmDocxTrans.exe llm=chatgpt apikey=sk-xxx input=./article.docx output=./translated.docx");
    // Console.WriteLine("  LlmDocxTrans.exe llm=deepseek apikey=xxx input=\"D:\\files\\paper.docx\" source=en target=zh-CN");
}