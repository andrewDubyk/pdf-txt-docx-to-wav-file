using System;
using System.Web;
using System.Speech.Synthesis;
using System.Threading.Tasks;
using iTextSharp.text.pdf;
using System.IO;
using System.Text;
using ICSharpCode.SharpZipLib.Zip;
using System.Xml;

namespace WebApplication1.Models
{
    public static class TypeConverter
    {
        public static async Task FileToWav(HttpPostedFileBase file, string path)
        {
            if (!File.Exists(path + file.FileName))
            {
                var extension = Path.GetExtension(file.FileName);
                var text = "";

                if (extension == ".pdf")
                {
                    using (PdfReader pfdReader = new PdfReader(file.InputStream))
                    {
                        for (var i = 1; i <= pfdReader.NumberOfPages; i++)
                        {
                            text += iTextSharp.text.pdf.parser.PdfTextExtractor.GetTextFromPage(pfdReader, i);
                        }
                    }
                }
                else if (extension == ".txt")
                {
                    using (var reader = new StreamReader(file.InputStream))
                    {
                        text = reader.ReadToEnd();
                    }
                }
                else if (extension == ".docx")
                {
                    file.SaveAs(path + file.FileName);
                    DocxToText dtt = new DocxToText(path + file.FileName);
                    text = dtt.ExtractText();
                    File.Delete(path + file.FileName);
                }

                await Task.Run(() =>
                {
                    using (var reader = new SpeechSynthesizer())
                    {
                        var fileName = path + Path.GetFileNameWithoutExtension(file.FileName) + ".wav";
                        reader.SelectVoiceByHints(VoiceGender.Female, VoiceAge.Senior);
                        reader.SetOutputToWaveFile(fileName);
                        var builder = new PromptBuilder();
                        builder.AppendText(text);
                        reader.Speak(builder);
                    }
                });
            }
        }

        public class DocxToText
        {
            private const string ContentTypeNamespace =
                @"http://schemas.openxmlformats.org/package/2006/content-types";

            private const string WordprocessingMlNamespace =
                @"http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            private const string DocumentXmlXPath =
                "/t:Types/t:Override[@ContentType=\"" +
        "application/vnd.openxmlformats-officedocument." +
        "wordprocessingml.document.main+xml\"]";

            private const string BodyXPath = "/w:document/w:body";

            private string docxFile = "";
            private string docxFileLocation = "";

            public DocxToText(string fileName)
            {
                docxFile = fileName;
            }

            #region ExtractText()
            public string ExtractText()
            {
                docxFileLocation = FindDocumentXmlLocation();
                return ReadDocumentXml();
            }
            #endregion

            #region FindDocumentXmlLocation()
            private string FindDocumentXmlLocation()
            {
                ZipFile zip = new ZipFile(docxFile);
                foreach (ZipEntry entry in zip)
                {

                    if (string.Compare(entry.Name, "[Content_Types].xml", true) == 0)
                    {
                        Stream contentTypes = zip.GetInputStream(entry);

                        XmlDocument xmlDoc = new XmlDocument
                        {
                            PreserveWhitespace = true
                        };
                        xmlDoc.Load(contentTypes);

                        contentTypes.Close();

                        XmlNamespaceManager nsmgr =
                            new XmlNamespaceManager(xmlDoc.NameTable);
                        nsmgr.AddNamespace("t", ContentTypeNamespace);

                        XmlNode node = xmlDoc.DocumentElement.SelectSingleNode(
                            DocumentXmlXPath, nsmgr);

                        if (node != null)
                        {
                            string location = ((XmlElement)node).GetAttribute("PartName");
                            zip.Close();
                            return location.TrimStart(new char[] { '/' });
                        }
                        break;
                    }
                }
                zip.Close();
                return null;
            }
            #endregion

            #region ReadDocumentXml()

            private string ReadDocumentXml()
            {
                StringBuilder sb = new StringBuilder();

                ZipFile zip = new ZipFile(docxFile);
                foreach (ZipEntry entry in zip)
                {
                    if (string.Compare(entry.Name, docxFileLocation, true) == 0)
                    {
                        Stream documentXml = zip.GetInputStream(entry);

                        XmlDocument xmlDoc = new XmlDocument
                        {
                            PreserveWhitespace = true
                        };
                        xmlDoc.Load(documentXml);
                        documentXml.Close();

                        XmlNamespaceManager nsmgr =
                            new XmlNamespaceManager(xmlDoc.NameTable);
                        nsmgr.AddNamespace("w", WordprocessingMlNamespace);

                        XmlNode node =
                            xmlDoc.DocumentElement.SelectSingleNode(BodyXPath, nsmgr);

                        if (node == null)
                        {
                            zip.Close();
                            return string.Empty;
                        }
                        sb.Append(ReadNode(node));

                        break;
                    }
                }
                zip.Close();
                return sb.ToString();
            }
            #endregion

            #region ReadNode()
            private string ReadNode(XmlNode node)
            {
                if (node == null || node.NodeType != XmlNodeType.Element)
                {
                    return string.Empty;
                }
                StringBuilder sb = new StringBuilder();
                foreach (XmlNode child in node.ChildNodes)
                {
                    if (child.NodeType != XmlNodeType.Element) continue;

                    switch (child.LocalName)
                    {
                        case "t":
                            sb.Append(child.InnerText.TrimEnd());

                            string space =
                                ((XmlElement)child).GetAttribute("xml:space");
                            if (!string.IsNullOrEmpty(space) &&
                                space == "preserve")
                                sb.Append(' ');
                            break;

                        case "cr":
                        case "br":
                            sb.Append(Environment.NewLine);
                            break;

                        case "tab":
                            sb.Append("\t");
                            break;

                        case "p":
                            sb.Append(ReadNode(child));
                            sb.Append(Environment.NewLine);
                            sb.Append(Environment.NewLine);
                            break;

                        default:
                            sb.Append(ReadNode(child));
                            break;
                    }
                }
                return sb.ToString();
            }
            #endregion
        }
    }
}