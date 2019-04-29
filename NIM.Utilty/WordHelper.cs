using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NIM.Utilty
{
    public static class WordHelper
    {

        public static void SearchAndReplaces(string wordLocalPath, List<(string replacedString, string replacingString)> items, Action<WordprocessingDocument, string> replacedCallback = null)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(wordLocalPath, true))
            {
                wordDoc.SearchAndReplaces(items, replacedCallback);
            }
        }

        public static void SearchAndReplaces(this WordprocessingDocument wordDoc, List<(string replacedString, string replacingString)> items,
            Action<WordprocessingDocument, string> replacedCallback = null)
        {

            string docText = null;
            using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
            {
                docText = sr.ReadToEnd();
            }
            items.ForEach(t =>
            {
                docText = docText.Replace(t.replacedString, t.replacingString);

            });

            replacedCallback?.Invoke(wordDoc, docText);

            using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
            {
                sw.Write(docText);
            }
            wordDoc.Save();

        }



        public static void SearchAndReplace(this WordprocessingDocument wordDoc, string replacedString, string replacingString)
        {
            string docText = null;
            using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
            {
                docText = sr.ReadToEnd();
            }

            // Regex regexText = new Regex(replacedString);
            docText = docText.Replace(replacedString, replacingString);
            // docText = regexText.Replace(docText, replacedString);

            using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
            {
                sw.Write(docText);
            }
        }


        public static void SaveDocument(this WordprocessingDocument wordDoc, string docText)
        {
            using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
            {
                sw.Write(docText);
            }
            wordDoc.Save();
        }

        public static string Replace(string docText, string replacedString, string replacingString)
        {
            docText = docText.Replace(replacedString, replacingString);
            return docText;
        }



        public static WordprocessingDocument OpenDocument(string wordLocalPath)
        {
            return WordprocessingDocument.Open(wordLocalPath, true);
        }
        public static void SearchAndReplace(string document, string replacedString, string replacingString)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
            {
                string docText = null;
                using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }

                // Regex regexText = new Regex(replacedString);
                docText = docText.Replace(replacedString, replacingString);
                // docText = regexText.Replace(docText, replacedString);

                using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                {
                    sw.Write(docText);
                }
                wordDoc.Save();
            }
        }




        public static void SetKeyWordItalicStyle(this Paragraph paragraph, string fullString, params char[] keys)
        {
            if (keys == null || keys.Length == 0)
                return;
            var originalRun = paragraph.Elements<Run>().First();
            var originalText = originalRun.Elements<Text>().First();

            var newText = (Text)originalText.Clone();
            newText.Text = "";
            var newRun = (Run)originalRun.Clone();
            originalRun.RemoveChild(originalText);
            paragraph.RemoveChild(originalRun);



            var lastWord = "";
            Text _text;
            Run _run;
            for (var i = 0; i < fullString.Length; i++)
            {
                var s = fullString[i];
                if (keys.Contains(s))
                {
                    if (lastWord != "")
                    {
                        _run = (Run)newRun.Clone();
                        _text = (Text)newText.Clone();
                        _text.Text = lastWord;
                        _run.Append(_text);
                        paragraph.Append(_run);
                    }


                    _text = (Text)newText.Clone();
                    _text.Text = s.ToString();
                    _run = (Run)originalRun.Clone();
                    _run.Append(_text);
                    paragraph.Append(_run);

                    var runProperties = _run.Elements<RunProperties>().FirstOrDefault();
                    if (runProperties == null)
                    {
                        runProperties = new RunProperties();
                        _run.Append(runProperties);
                    }
                    var italic = runProperties.Elements<Italic>().FirstOrDefault();
                    if (italic == null)
                    {
                        italic = new Italic();
                        runProperties.Append(italic);
                    }
                    lastWord = "";
                }

                else
                    lastWord += s;
            }
            if (lastWord != "")
            {
                _run = (Run)originalRun.Clone();
                _text = (Text)newText.Clone();
                _text.Text = lastWord;
                _run.Append(_text);
                paragraph.Append(_run);

            }

        }



        public static void SetNumberStyle(this WordprocessingDocument document, string fullString)
        {
            //var keys = new char[] { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' };
            //if (keys == null || keys.Length == 0)
            //    return;
            var originalTexts = document.MainDocumentPart.Document.Body.Descendants<Text>().Where(t => t.Text == fullString).ToList();
            originalTexts.ForEach(originalText =>
            {
                originalText.Text = "";

                var originalRun = (Run)originalText.Parent;

                var newText = (Text)originalText.Clone();
                newText.Text = "";
                var newRun = (Run)originalRun.Clone();

                var paragraph = (Paragraph)originalRun.Parent;
                originalRun.RemoveChild(originalText);
                paragraph.RemoveChild(originalRun);



                var lastWord = "";
                Text _text;
                Run _run;
                for (var i = 0; i < fullString.Length; i++)
                {
                    var s = fullString[i];
                    if (!NumberHelper.IsChinese(s.ToString()))
                    {
                        if (lastWord != "")
                        {
                            _run = (Run)newRun.Clone();
                            _text = (Text)newText.Clone();
                            _text.Text = lastWord;
                            _run.Append(_text);
                            paragraph.Append(_run);
                        }


                        _text = (Text)newText.Clone();
                        _text.Text = s.ToString();
                        _run = (Run)originalRun.Clone();
                        _run.Append(_text);
                        paragraph.Append(_run);

                        var runProperties = _run.Elements<RunProperties>().FirstOrDefault();
                        if (runProperties == null)
                        {
                            runProperties = new RunProperties();
                            _run.Append(runProperties);
                        }
                        var runFonts = runProperties.Elements<RunFonts>().FirstOrDefault();
                        if (runFonts == null)
                        {
                            runFonts = new RunFonts();
                            runProperties.Append(runFonts);
                        }
                        runFonts.Ascii = "Times New Roman";
                        runFonts.HighAnsi = "Times New Roman";
                        runFonts.ComplexScript = "Times New Roman";
                        runFonts.EastAsia = "Times New Roman";
                        lastWord = "";
                    }

                    else
                        lastWord += s;
                }
                if (lastWord != "")
                {
                    _run = (Run)originalRun.Clone();
                    _text = (Text)newText.Clone();
                    _text.Text = lastWord;
                    _run.Append(_text);
                    paragraph.Append(_run);

                }

            });

        }

        /// <summary>
        /// 设置复杂字体，包括斜体，和非中文字体的times new roman 字休
        /// </summary>
        /// <param name="document"></param>
        /// <param name="fullString"></param>
        /// <param name="values"></param>
        public static void ComplexReplace(this Paragraph paragraph, string fullString, params char[] keys)
        {

            if (keys == null)
                keys = new char[] { };

            var originalRun = paragraph.Elements<Run>().First();
            var originalText = originalRun.Elements<Text>().First();

            var newText = (Text)originalText.Clone();
            newText.Text = "";
            var newRun = (Run)originalRun.Clone();
            originalRun.RemoveChild(originalText);
            paragraph.RemoveChild(originalRun);



            var lastWord = "";
            Text _text;
            Run _run;
            for (var i = 0; i < fullString.Length; i++)
            {
                var s = fullString[i];
                if (keys.Contains(s) || !NumberHelper.IsChinese(s.ToString()))
                {

                    if (lastWord != "")
                    {
                        _run = (Run)newRun.Clone();
                        _text = (Text)newText.Clone();
                        _text.Text = lastWord;
                        _run.Append(_text);
                        paragraph.Append(_run);
                    }


                    _text = (Text)newText.Clone();
                    _text.Text = s.ToString();
                    _run = (Run)originalRun.Clone();
                    _run.Append(_text);
                    paragraph.Append(_run);

                    var runProperties = _run.Elements<RunProperties>().FirstOrDefault();
                    if (runProperties == null)
                    {
                        runProperties = new RunProperties();
                        _run.Append(runProperties);
                    }
                    if (keys.Contains(s))
                    {
                        var italic = runProperties.Elements<Italic>().FirstOrDefault();
                        if (italic == null)
                        {
                            italic = new Italic();
                            runProperties.Append(italic);
                        }
                    }
                    else
                    {
                        var runFonts = runProperties.Elements<RunFonts>().FirstOrDefault();
                        if (runFonts == null)
                        {
                            runFonts = new RunFonts();
                            runProperties.Append(runFonts);
                        }
                        runFonts.Ascii = "Times New Roman";
                        runFonts.HighAnsi = "Times New Roman";
                        runFonts.ComplexScript = "Times New Roman";
                        runFonts.EastAsia = "Times New Roman";
                    }
                    lastWord = "";
                }

                else
                    lastWord += s;
            }
            if (lastWord != "")
            {
                _run = (Run)originalRun.Clone();
                _text = (Text)newText.Clone();
                _text.Text = lastWord;
                _run.Append(_text);
                paragraph.Append(_run);

            }

        }




        private static void _SetCompltexReplace(Paragraph paragraph, string value, Action<Paragraph, string> replaceAction = null)
        {
            if (replaceAction == null)
                paragraph.Elements<Run>().First().Elements<Text>().First().Text = value;
            else
                replaceAction(paragraph, value);



        }
    }

}
