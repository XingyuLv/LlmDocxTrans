using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using LLmDocxTrans;


public class LlmDocxTrans
{
    private readonly string _docxFilePath;

    public LlmDocxTrans(string docxFilePath)
    {
        _docxFilePath = docxFilePath;
    }

    private string OutputDocxFilePath => _docxFilePath[.._docxFilePath.LastIndexOf(".", StringComparison.Ordinal)] + "-llmResult.docx";

    private string OutputXmlDir
    {
        get
        {
            var dir = _docxFilePath[.._docxFilePath.LastIndexOf(".", StringComparison.Ordinal)] + "-llmXml";
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            return dir;
        }
    }


    public static void Process(string inputDocx, string llmType, string apikey)
    {
        var llmDocxTrans = new LlmDocxTrans(inputDocx);
        llmDocxTrans.ExtractParagraphsToXml();
        ChatClient chatClient = null!;
        switch (llmType)
        {
            case "deepseek":
                chatClient = new DeepSeekChatClient(apikey);
                break;
            case "doubao":
                chatClient = new DoubaoChatClient(apikey);
                break;
            case "chatgpt":
                break;
        }
        llmDocxTrans.LlmProcess(llmType, chatClient, llmDocxTrans.OutputXmlDir);
        llmDocxTrans.GenerateDocxFromXmlFiles();
    }

    public void LlmProcess(string llm, ChatClient chatClient,  string outputXmlDir)
    {
        var translatedXmlFiles = Directory.GetFiles(outputXmlDir, "paragraph_*.xml")
            .Select(filePath =>
            {
                var fileName = Path.GetFileNameWithoutExtension(filePath);
                var number = int.Parse(fileName.Split('_')[1]);
                return new { FilePath = filePath, Number = number };
            })
            .OrderBy(x => x.Number)
            .Select(x => x.FilePath)
            .Where(x => !x.Contains("transResult"))
            .ToList();

        foreach (var originXmlPath in translatedXmlFiles)
        {
            var originXml = File.ReadAllText(originXmlPath);
            var transXmlPath = originXmlPath.Replace(".xml", "_transResult.xml");

            // 构建prompt
            string prompt = @"按照如下规则翻译我所给你的xml：
1、为保证翻译质量，你需要将所有<r>标签内的文本拼接成一段话进行翻译，我会找专业人士对你的翻译结果进行质量评价
2、<r>标签中的""_rid""属性代表当前标签内文本携带的单独样式，你只能使用原始待翻译xml中已存在的""_rid""，绝对不能编造不存在的""_rid""
3、如果某个词单独存放在一个<r>标签内，你要注意这个词是可能有单独样式的
4、为保证翻译质量，你可以增加、减少或合并<r>标签，但注意你要为<r>标签选择合适的""_rid""样式
5、返回的译文也需要携带标签，格式与原始xml一致，并且要保证标签完整性
6、只返回结果即可，无需解释
7、en2zh方向，原始待翻译xml如下：
" + originXml;

            string transResult = string.Empty;
            int retryCount = 0;
            bool success = true;

            try
            {
                do
                {
                    var start = DateTime.Now;
                    transResult = chatClient.Chat(prompt);
                    Console.WriteLine($"{transXmlPath}翻译完成，耗时{(DateTime.Now - start).TotalSeconds}s({llm})");
                    // 校验返回的XML是否合法
                    var errorMsg = XmlValidation.IsValidXml(originXml, transResult);
                    if (errorMsg != null)
                    {
                        success = false;
                        retryCount++;
                        prompt += @$"
你上次输出翻译的结果如下：
{transResult}
存在如下错误：
{errorMsg}
请重新按照要求完成翻译任务！
";
                        Console.WriteLine(@$"{transXmlPath}校验未通过，新的prompt如下：
{prompt}"
                        );
                    }
                    else
                    {
                        Console.WriteLine($"{transXmlPath}校验通过!");
                    }
                } while (!success && retryCount < 2);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"API调用出错: {ex.Message}");
                throw;
            }


            if (success)
            {
                // 保存翻译结果
                File.WriteAllText(transXmlPath, transResult);
                // 
            }
            else
            {
                Console.WriteLine($"翻译失败: {originXmlPath}");
                throw new Exception();
            }
        }
    }

    /// <summary>
    /// 将docx文件的每个段落解析成单独的XML文件
    /// </summary>
    private void ExtractParagraphsToXml()
    {
        // 打开docx文件
        using (WordprocessingDocument doc = WordprocessingDocument.Open(_docxFilePath, true))
        {
            // 获取文档主体部分
            Document? document = doc.MainDocumentPart?.Document;

            // 获取所有段落
            var paragraphs = document?.Descendants<Paragraph>().ToList();
            for (var i = 0; i < paragraphs?.Count; i++)
            {
                var paragraph = paragraphs[i];

                // 创建编号
                paragraph.CreateRunId();

                // 为每个段落创建XML文件
                var xmlFilePath = Path.Combine(OutputXmlDir, $"paragraph_{i}.xml");
                // 使用 PowerTools 简化 XML  不存在超链
                var xmlParagraph = new Paragraph();
                foreach (var ele in paragraph.Elements())
                {
                    if (ele is Run run)
                    {
                        xmlParagraph.AppendChild(run.CloneNode(true));
                    }

                    if (ele is Hyperlink hyperlink)
                    {
                        foreach (var hrun in hyperlink.Elements<Run>())
                        {
                            xmlParagraph.AppendChild(hrun.CloneNode(true));
                        }
                    }
                }

                string cleanedXml = CleanXml(xmlParagraph.OuterXml);
                // 直接将段落的OuterXml写入文件
                File.WriteAllText(xmlFilePath, cleanedXml);
            }
        }
    }

    private void GenerateDocxFromXmlFiles()
    {
        var translatedXmlFiles = Directory.GetFiles(OutputXmlDir, "paragraph_*.xml")
            .Select(filePath =>
            {
                var fileName = Path.GetFileNameWithoutExtension(filePath);
                var number = int.Parse(fileName.Split('_')[1]);
                return new { FilePath = filePath, Number = number };
            })
            .OrderBy(x => x.Number)
            .Select(x => x.FilePath)
            .Where(x => x.Contains("transResult"))
            .ToList();
        try
        {
            // 复制原始文件作为新文件的基础
            File.Copy(_docxFilePath, OutputDocxFilePath, true);

            // 打开新文件进行编辑
            // using var originDocx = WordprocessingDocument.Open(DocxFilePath, true);
            using var outputDocx = WordprocessingDocument.Open(OutputDocxFilePath, true);
            // var originDocument = originDocx.MainDocumentPart!.Document;
            var outputDocument = outputDocx.MainDocumentPart!.Document;

            // 获取原始文件中的所有段落
            var outputParagraphList = outputDocument.Descendants<Paragraph>().ToList();
            // var originParagraphDict = originDocument.Descendants<Paragraph>()
            //     .ToDictionary(c => c.GetIndex(), c => c);

            for (var i = 0; i < outputParagraphList.Count; i++)
            {
                var outputParagraph = outputParagraphList[i];

                // 跳过空段落
                if (outputParagraph.InnerText.Trim().Length == 0)
                {
                    continue;
                }

                // 读取原始run clone备用
                var originRunDict = new Dictionary<string, OpenXmlElement>();
                var elementForAppend = new List<OpenXmlElement>();
                var elementForRemove = new List<OpenXmlElement>();
                var hyperlinkMapping = outputParagraph.CreateHyperlinkMapping();

                foreach (var ele in outputParagraph.Elements())
                {
                    if (ele is Run run)
                    {
                        originRunDict.Add(run.GetRunId(), run);
                        elementForRemove.Add(run);
                    }

                    if (ele is Hyperlink hyperlink)
                    {
                        foreach (var hrun in hyperlink.Elements<Run>())
                        {
                            originRunDict.Add(hrun.GetRunId(), hrun);
                        }

                        elementForRemove.Add(hyperlink);
                    }
                }

                // 读取翻译后的XML，逐一写回run hy
                XElement? xParagraph = XDocument.Parse(File.ReadAllText(translatedXmlFiles[i])).Root;
                List<XElement>? xElements = xParagraph?.Elements().ToList()!;
                for (var j = 0; j < xElements.Count; j++)
                {
                    XElement xRun = xElements[j];
                    var originRun = originRunDict.GetValueOrDefault(xRun.GetRunId());

                    // 空run
                    try
                    {
                        try
                        {
                            if (originRun.Parent is Paragraph paragraph)
                            {
                                var newRun = originRun!.CloneNode(true);
                                if (!string.IsNullOrEmpty(string.Join("",
                                        newRun.Elements<Text>().Select(t => t.Text).ToList())))
                                {
                                    newRun.Elements<Text>().ElementAt(0).Text = xRun.Elements().ElementAt(0).Value;
                                    newRun.Elements<Text>().ElementAt(0).Space = SpaceProcessingModeValues.Preserve;
                                }

                                elementForAppend.Add(newRun);
                            }
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e);
                            throw;
                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e);
                        throw;
                    }

                    // 空hy
                    if (originRun.Parent is Hyperlink hyperlink)
                    {
                        // originRun是否同属于一个hyperlink:
                        var newRun = originRun!.CloneNode(true);
                        if (!string.IsNullOrEmpty(string.Join("",
                                newRun.Elements<Text>().Select(t => t.Text).ToList())))
                        {
                            newRun.Elements<Text>().ElementAt(0).Text = xRun.Elements().ElementAt(0).Value;
                            newRun.Elements<Text>().ElementAt(0).Space = SpaceProcessingModeValues.Preserve;
                        }

                        var belongHyperLink = originRun.GetBelongHyperLink(hyperlinkMapping);
                        belongHyperLink.AppendChild(newRun);

                        // 下一个的father不是hylink，退出
                        if (j + 1 <= xElements.Count - 1)
                        {
                            if (originRunDict.GetValueOrDefault(xElements[j + 1].GetRunId())!.Parent is Paragraph)
                            {
                                elementForAppend.Add(belongHyperLink);
                            }
                        }

                        // 最后一个hylink
                        if (j == xElements.Count - 1)
                        {
                            elementForAppend.Add(belongHyperLink);
                        }
                    }
                }

                // hyperlink


                // 执行清除
                elementForRemove.ForEach(element => element.Remove());

                // 执行添加
                elementForAppend.ForEach(element => outputParagraph.AppendChild(element));

                // 写入output p

                // 替换原始段落
                // Paragraph newParagraph = new Paragraph(newParagraphXml);
                //
                // originalParagraph.Parent.ReplaceChild(newParagraph, originalParagraph);
            }

            // 保存文档
            outputDocument.Save();
            Console.WriteLine($"译文生成成功：{OutputDocxFilePath}");

        }
        catch (Exception ex)
        {
            Console.WriteLine($"生成翻译后的docx文件时发生错误: {ex.Message}");
            Console.WriteLine(ex);
        }
    }


    /// <summary>
    /// 简化XML结构，只保留标签和文本内容
    /// </summary>
    private string CleanXml(string xml)
    {
        XDocument doc = XDocument.Parse(xml);

        // 递归清理所有元素
        CleanElement(doc.Root);

        return doc.ToString();
    }

    private void CleanElement(XElement? element)
    {
        if (element == null)
        {
            return;
        }

        // 保留包含_number的属性
        element.Attributes()
            .Where(attr => !attr.Name.LocalName.Contains("_rid"))
            .Remove();

        // 创建新的无命名空间的名称
        element.Name = element.Name.LocalName;

        // 只保留 w:p, w:r, w:t 这些核心标签
        var children = element.Elements().ToList();
        foreach (var child in children)
        {
            var localName = child.Name.LocalName;
            if (localName is "p" or "r" or "t")
            {
                CleanElement(child);
            }
            else
            {
                child.Remove();
            }
        }
    }
}