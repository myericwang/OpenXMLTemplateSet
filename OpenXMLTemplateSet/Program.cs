using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace DocxTemplateSet
{
    class Program
    {
        static void Main(string[] args)
        {
            string templateName = "DocxTemplate";
            StreamReader r = new StreamReader(@"SourceData\test.json");
            string jsonString = r.ReadToEnd();
            Dictionary<string, object> data = JsonConvert.DeserializeObject<Dictionary<string, object>>(jsonString);


            var docxBytes = WordRender.GenerateDocx(File.ReadAllBytes($@"template\{templateName}.docx"), data);
            if (!Directory.Exists(@"D:\OpenXml_rep"))
            {
                Directory.CreateDirectory(@"D:\OpenXml_rep");
            }
            File.WriteAllBytes($@"D:\OpenXml_rep\report_{templateName}-{DateTime.Now:yyMMddHHmmss}.docx", docxBytes);

        }
    }

    public static class WordRender
    {
        static void ReplaceParserTag(this OpenXmlElement elem, Dictionary<string, dynamic> datas)
        {
            XmlMutiTransform(elem, datas, string.Empty);
            XmlSingleTransform(elem, datas);
        }


        private static void XmlSingleTransform(OpenXmlElement elem, Dictionary<string, dynamic> data)
        {
            var run_pool = new List<Run>();
            var poolS = new List<Run>();
            var innerTextString = string.Empty;
            var matchTextS = string.Empty;

            //找出鮮明提示
            var hiliteRuns = elem.Descendants<Run>().Where(o => o.RunProperties?.Elements<Highlight>().Any() ?? false).ToList();
            foreach (var run in hiliteRuns)
            {
                string text = run.InnerText;

                if (text.StartsWith("$"))
                {
                    poolS = new List<Run> { run };
                    matchTextS = text;
                }
                else
                {
                    matchTextS += text;
                    poolS.Add(run);
                }
                if (text.EndsWith("}"))
                {
                    var m = Regex.Match(matchTextS, @"\$\{(?<n>\w+)\}");
                    if (m.Success && data.ContainsKey(m.Groups["n"].Value))
                    {
                        var firstRun = poolS.First();
                        firstRun.RemoveAllChildren<Text>();
                        firstRun.RunProperties.RemoveAllChildren<Highlight>();

                        var newText = data[m.Groups["n"].Value].ToString();
                        var firstLine = true;
                        foreach (var line in Regex.Split(newText, @"\\n"))
                        {
                            if (firstLine) firstLine = false;
                            else firstRun.Append(new Break());
                            firstRun.Append(new Text(line));
                        }
                        poolS.Skip(1).ToList().ForEach(o => o.Remove());
                    }
                }
            }
        }

        private static void XmlMutiTransform(OpenXmlElement elem, Dictionary<string, dynamic> datas, string layer)
        {
            string matchText = string.Empty;
            int start = -1;
            for (int i = 0; i < elem.ChildElements.Count; i++)
            {
                if (elem.ChildElements[i].InnerText.StartsWith("#"))
                {
                    start = i;
                }
                else if (start == -1 && elem.ChildElements[i].InnerText.Contains("#"))
                {
                    foreach(var child in elem.ChildElements[i])
                    {
                        XmlMutiTransform(child, datas, layer);
                    }
                }

                if (start >= 0)
                {
                    matchText += elem.ChildElements[i].InnerText;

                    if (elem.ChildElements[i].InnerText.EndsWith($"]{layer}"))
                    {
                        var m = Regex.Match(matchText, $@"(?<=\#)\S+(?=\]{layer})");
                        var matchkey = Regex.Match(m.Value, @"\S+?(?=\[)");

                        if (datas.ContainsKey(matchkey.Groups[0].Value))
                        {
                            int row = 0;
                            elem.ChildElements[i].Remove();
                            foreach (var data in datas[matchkey.Groups[0].Value])
                            {
                                Dictionary<string, dynamic> cols = JsonConvert.DeserializeObject<Dictionary<string, dynamic>>(JsonConvert.SerializeObject(data));

                                for (int k = 0; k < i - start - 1; k++)
                                {
                                    elem.InsertAt(elem.ChildElements[start + k + 1].CloneNode(true), i + k + ((i - start - 1) * row));

                                    var l = Regex.Match(m.Value, $@"(?<=\#)\S+(?=\]{layer})");
                                    var l_matchkey = Regex.Match(l.Value, @"\w+(?=\[)");
                                    if (l.Success)
                                    {
                                        Dictionary<string, dynamic> tmp = new Dictionary<string, dynamic>();
                                        tmp.Add(l_matchkey.Value, cols[l_matchkey.Value]);

                                        XmlMutiTransform(elem.ChildElements[i + k + ((i - start - 1) * row)], cols, $@"{layer}.");
                                    }
                                    XmlSingleTransform(elem.ChildElements[i + k + ((i - start - 1) * row)], cols);
                                }
                                row++;
                            }
                            for (int k = 0; k < i - start; k++)
                            {
                                elem.ChildElements[start].Remove();
                            }
                            matchText = string.Empty;
                        }
                    }
                }
            }
        }

        public static byte[] GenerateDocx(byte[] template, Dictionary<string, object> data)
        {
            using (var ms = new MemoryStream())
            {
                ms.Write(template, 0, template.Length);
                using (var docx = WordprocessingDocument.Open(ms, true))
                {
                    docx.MainDocumentPart.HeaderParts.ToList().ForEach(hdr =>
                    {
                        hdr.Header.ReplaceParserTag(data);
                    });
                    docx.MainDocumentPart.FooterParts.ToList().ForEach(ftr =>
                    {
                        ftr.Footer.ReplaceParserTag(data);
                    });
                    docx.MainDocumentPart.Document.Body.ReplaceParserTag(data);
                    docx.Save();
                }
                return ms.ToArray();
            }
        }
    }


    public static class Trans
    {
        public static Dictionary<string, T> ToDictionary<T>(object obj)
        {
            return JsonConvert.DeserializeObject<Dictionary<string, T>>(JsonConvert.SerializeObject(obj));
        }
    }
}
