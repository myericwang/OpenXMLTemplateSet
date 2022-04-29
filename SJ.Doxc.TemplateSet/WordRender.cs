using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;
using SJConvert;
using System.Text.RegularExpressions;

namespace SJ.Doxc.TemplateSet
{
    public class DocTemplateSet
    {
        /// <summary>
        /// docx範本轉為bytes進行處理，並將檔案依指定檔名(套件會自動加上.docx)儲存於指定的資料夾中
        /// </summary>
        /// <param name="templateBytes">word範本的bytes[]</param>
        /// <param name="data">參數物件</param>
        /// <param name="outputFolder">檔案需儲存的路徑</param>
        /// <returns>檔案儲存的路徑</returns>
        public string StartUp(byte[] templateBytes, object data, string outputFolder, string fileName)
        {
            var docxBytes = WordRender.GenerateDocx(templateBytes, DictionaryEx.ToDictionary<dynamic>(data));
            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }

            string outputFilePath = $@"{outputFolder}\{fileName}.docx";
            File.WriteAllBytes(outputFilePath, docxBytes);

            return outputFilePath;
        }

        /// <summary>
        /// docx範本轉為bytes進行處理，並用bytes[]回傳結果
        /// </summary>
        /// <param name="templateBytes">word範本的bytes[]</param>
        /// <param name="data">參數物件</param>
        /// <returns>byte[]的處理結果</returns>
        public byte[] StartUp(byte[] templateBytes, object data)
        {
            return WordRender.GenerateDocx(templateBytes, DictionaryEx.ToDictionary<dynamic>(data));
        }

        /// <summary>
        /// 指定docx範本路徑進行處理，並用bytes[]回傳結果
        /// </summary>
        /// <param name="templatePath">word範本的路徑</param>
        /// <param name="data">參數物件</param>
        /// <returns></returns>
        public byte[] StartUp(string templatePath, object data)
        {
            var templateBytes = File.ReadAllBytes(templatePath);
            return WordRender.GenerateDocx(templateBytes, DictionaryEx.ToDictionary<dynamic>(data));
        }

        /// <summary>
        /// 指定docx範本路徑轉進行處理，並將檔案依指定檔名(套件會自動加上.docx)儲存於指定的資料夾中
        /// </summary>
        /// <param name="templateBytes">word範本的bytes[]</param>
        /// <param name="data">參數物件</param>
        /// <param name="outputFolder">檔案需儲存的路徑</param>
        /// <returns>檔案儲存的路徑</returns>
        public string StartUp(string templatePath, object data, string outputFolder, string fileName)
        {
            var templateBytes = File.ReadAllBytes(templatePath);
            return StartUp(templateBytes, data, outputFolder, fileName);
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
                    foreach (var child in elem.ChildElements[i])
                    {
                        XmlMutiTransform(child, datas, layer);
                    }
                }

                if (start >= 0)
                {
                    matchText += elem.ChildElements[i].InnerText;

                    if (elem.ChildElements[i].InnerText.EndsWith($"]{layer}"))
                    {
                        var m = Regex.Match(matchText, $@"(?<=\#)[\S\s]+(?=\]{layer})");
                        var matchkey = Regex.Match(m.Value, @"[\S\s]+?(?=\[)");

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

                                    var l = Regex.Match(m.Value, $@"(?<=\#)[\S\s]+(?=\]{layer})");
                                    var l_matchkey = Regex.Match(l.Value, @"[\S\s]+(?=\[)");
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
}