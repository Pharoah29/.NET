using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Web;

using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.tool.xml.css;
using iTextSharp.tool.xml;
using iTextSharp.tool.xml.html;
using iTextSharp.tool.xml.pipeline.html;
using iTextSharp.tool.xml.pipeline.end;
using iTextSharp.tool.xml.pipeline.css;
using iTextSharp.tool.xml.parser;
using iTextSharp.text.html.simpleparser;

namespace Utils.Classes.PDF
{
    /// <summary>
    /// Converts html or plain text to a PDF stream,
    /// handles hebrew (UTF-8) formatting
    /// </summary>
    /// <example>
    /// PDFOutput pdf = PDFConverter.LoadFiles(files);
    /// PDFOutput pdf = PDFConverter.LoadFromFile(file);
    /// PDFOutput pdf = PDFConverter.LoadText(html)
    /// PDFConverter.Render(pdf); 
    /// </example> 
    public class PDFConverter
    {

        /// <summary>
        /// Load an HTML file to convert
        /// </summary>
        /// <param name="filePath">Specify the file path</param>
        /// <returns></returns>
        public static PDFOutput LoadFromFile(string filePath)
        {

            StreamReader srPDF = new StreamReader(filePath, Encoding.UTF8);

            string text = srPDF.ReadToEnd();

            srPDF.Close();

            string htmlText = CreateHTMLBody(text);

            PDFOutput output = new PDFOutput
            {
                FileName = DateTime.Now.Ticks + ".pdf",

                Data = Convert(htmlText)
            };


           return output;
       }

        /// <summary>
        /// Load an array of HTML files to convert
        /// </summary>
        /// <param name="filePaths">Specify an array of files to convert</param>
        /// <returns>The binary data of the combined files</returns>
        public static PDFOutput LoadFiles(string[] filePaths)
        {

            StringBuilder sbHtml = new StringBuilder();

            foreach (string file in filePaths)
            {
                StreamReader srPDF = new StreamReader(file, Encoding.UTF8);

                string text = srPDF.ReadToEnd();

                srPDF.Close();

                sbHtml.Append(CreateHTMLBody(text));
            }


            PDFOutput output = new PDFOutput
            {
                FileName = DateTime.Now.Ticks + ".pdf",

                Data = Convert(sbHtml.ToString())
            };


            return output;

        }

        /// <summary>
        /// Load a text to convert. can be an HTML text as well
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public static PDFOutput LoadText(string text)
        {
            string htmlText = CreateHTMLBody(text);

            PDFOutput output = new PDFOutput
            {
                FileName = DateTime.Now.Ticks + ".pdf",

                Data = Convert(htmlText)
            };

            return output;

        }

        /// <summary>
        ///  //Render the HTML to a PDF file, and output it to the response stream
        /// </summary>
        /// <param name="data"></param>
        /// <param name="outputFilename"></param>
        public static void Render(PDFOutput pdf)
        {
           
            HttpContext.Current.Response.ClearContent();
            HttpContext.Current.Response.ClearHeaders();
            HttpContext.Current.Response.ContentType = "application/pdf";
            HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment; filename=" + pdf.FileName);

            HttpContext.Current.Response.BinaryWrite(pdf.Data);
            HttpContext.Current.Response.End();
            HttpContext.Current.Response.Flush();
            HttpContext.Current.Response.Clear();
        }


        private static string CreateHTMLBody(string html)
        {

            StringBuilder sbHtml = new StringBuilder(@"<div dir=""rtl"" style=""font-family: arial;"">");

            sbHtml.Append(html);

            sbHtml.Append("</div>");

            return sbHtml.ToString();
        }

        /// <summary>
        /// The actual converting using iTextSharp objects
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        private static byte[] Convert(string text)
        {
            Document doc = new Document(PageSize.LETTER, 50, 50, 50, 50);

            using (MemoryStream output = new MemoryStream())
            {
                PdfWriter writer = PdfWriter.GetInstance(doc, output);
                writer.RunDirection = 3;
                doc.Open();

                //Full path to the Unicode Arial file
                string ARIALUNI_TFF = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "ARIAL.TTF");

                //Create a base font object making sure to specify IDENTITY-H
                BaseFont bf = BaseFont.CreateFont(ARIALUNI_TFF, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);

                //Create a specific font object
                Font f = new Font(bf, 12, Font.NORMAL);

                var cssResolver = new StyleAttrCSSResolver();
                var fontProvider = new XMLWorkerFontProvider(XMLWorkerFontProvider.DONTLOOKFORFONTS);

                var cssAppliers = new CssAppliersImpl(fontProvider);
                var htmlContext = new HtmlPipelineContext(cssAppliers);
                htmlContext.SetTagFactory(Tags.GetHtmlTagProcessorFactory());

                var pdf = new PdfWriterPipeline(doc, writer);
                var html = new HtmlPipeline(htmlContext, pdf);
                var css = new CssResolverPipeline(cssResolver, html);

                var worker = new XMLWorker(css, true);
                var p = new XMLParser(worker);

                fontProvider.Register(ARIALUNI_TFF);

                using (var ms = new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes(text)))
                {
                    using (var sr = new StreamReader(ms))
                    {
                        p.Parse(sr);
                    }
                }

                doc.Close();

                byte[] arrOutput = output.ToArray();

                output.Close();

                return arrOutput;

            }

        }

    }

    public class PDFOutput
    {
        public string FileName { get; set; }

        public byte[] Data { get; set; }
    }
}
