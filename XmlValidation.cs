using System.Xml.Linq;

namespace LLmDocxTrans;

public static class XmlValidation
{
    
    /// <summary>
    /// 检查XML是否合法
    /// </summary>
    public static string? IsValidXml(string originXml, string transResultXml)
    {
        try
        {
            // 尝试解析XML
            XDocument.Parse(transResultXml);
        }
        catch
        {
            // Console.WriteLine($"xml不完整，originXml={originXml}");
            // Console.WriteLine($"xml不完整，transResultXml={transResultXml}");
            return "xml不完整！";
        }

        var idSet = new HashSet<string>();
        var xElement = XDocument.Parse(originXml).Root;
        foreach (var element in xElement?.Descendants().ToList()!)
        {
            var value = element.Attribute("_rid")?.Value;
            if (value != null)
            {
                idSet.Add(value);
            }
        }

        // 校验
        foreach (var element in XDocument.Parse(transResultXml).Root?.Descendants().ToList()!)
        {
            var value = element.Attribute("_rid")?.Value;
            if (value != null)
            {
                if (!idSet.Contains(value))
                {
                    // 错误！
                    var msg = @$"原始待翻译xml中不存在{"_rid"}=""{value}""！记住你只能使用原始待翻译xml中已存在的""_rid""，不能编造不存在""_rid""";
                    // Console.WriteLine(msg);
                    return msg;
                }
            }
        }


        return null;
    }
}