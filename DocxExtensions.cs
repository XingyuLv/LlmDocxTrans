using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using NuGet.Packaging;

namespace LLmDocxTrans;

public static class DocxExtensions
{
    private static int _index;
    private static int Index => _index++;


    public static string GetRunId(this OpenXmlElement ele)
    {
        // 创建标识
        try
        {
            var attribute = ele.GetAttribute("_rid", "");
            return attribute.Value!;
        }
        catch (Exception e)
        {
            Console.WriteLine(e.Message);
            throw;
        }
    }

    public static string GetRunId(this XElement xEle)
    {
        return xEle.GetOptionalAttributeValue("_rid");
    }

    public static void CreateRunId(this OpenXmlElement ele)
    {
        var allXmlList = ele.Descendants().ToList();
        foreach (var openXmlElement in allXmlList)
        {
            if (openXmlElement is Run run)
            {
                // 创建标识
                OpenXmlAttribute attribute;
                try
                {
                    attribute = run.GetAttribute("_rid", "");
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                    attribute = new OpenXmlAttribute
                    {
                        LocalName = "_rid",
                        Value = Index.ToString()
                    };
                    run.SetAttribute(attribute);
                }

                try
                {
                    var pNumber = int.Parse(attribute.Value!);
                    // Console.WriteLine($"create _rid={pNumber}");
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                    throw;
                }
            }
        }
    }
    
    public static Dictionary<OpenXmlElement, HashSet<string>> CreateHyperlinkMapping(this Paragraph paragraph)
    {
        var hyperlinkMapping = new Dictionary<OpenXmlElement, HashSet<string>>();
        foreach (var hyperlink in paragraph.Elements<Hyperlink>())
        {
            var idSet = new HashSet<string>();
            foreach (var run in hyperlink.Elements<Run>())
            {
                idSet.Add(run.GetRunId());
            }

            var openXmlElement = hyperlink.CloneNode(true);
            openXmlElement.Elements().ToList().ForEach(ele => ele.Remove());
            hyperlinkMapping.Add(openXmlElement.CloneNode(true), idSet);
        }

        return hyperlinkMapping;
    }
    
    public static OpenXmlElement GetBelongHyperLink(this OpenXmlElement run,
        Dictionary<OpenXmlElement, HashSet<string>> hyperlinkMapping)
    {
        foreach (var (belongHyperLink, idSet) in hyperlinkMapping)
        {
            if (idSet.Contains(run.GetRunId()))
            {
                return belongHyperLink;
            }
        }

        throw new Exception("no BelongHyperLink!");
    }

}