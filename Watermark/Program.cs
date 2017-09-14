using System;
using Spire.Doc;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Spire.Doc.Documents;
using Spire.Pdf;
using Spire.Pdf.Graphics;
using Spire.Presentation;
using Spire.Presentation.Drawing;
using FileFormat = Spire.Presentation.FileFormat;

namespace Watermark
{
    internal class Program
    {
         public static void Main(string[] args)
        {
            string fileDir = Environment.CurrentDirectory;

            string pic = "";
            string text = args.Length == 0 ? "ECNU" : args[0];

            string outPath = fileDir + "\\watermark\\";
            if (!Directory.Exists(outPath))
                Directory.CreateDirectory(outPath);

            try
            {
                //LicenseHelper.ModifyInMemory_Spire.ActivateMemoryPatching();
                DirectoryInfo dir = new DirectoryInfo(fileDir);
                foreach (FileInfo file in dir.GetFiles())
                {
                    switch (file.Extension)
                    {
                        case ".docx":
                            docWaterMark(file.FullName, outPath + file.Name, pic, text);
                            break;
                        case ".pptx":
                            pptWaterMark(file.FullName, outPath + file.Name, pic, text);
                            break;
                        case ".pdf":
                            pdfWaterMark(file.FullName, outPath + file.Name, pic, text);
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

        }

        public static void docWaterMark(String file, String output, String pic, String text)
        {
            Document doc = new Document();
            doc.LoadFromFile(file);

            if (pic != "")
            {
                //设图片水印
                PictureWatermark picture = new PictureWatermark();
                picture.Picture = Image.FromFile(pic);
                picture.Scaling = 100;
                picture.IsWashout = false;
                doc.Watermark = picture;
            }
            else
            {
                //设置文本水印
                TextWatermark txtWatermark = new TextWatermark();
                txtWatermark.Text = text;
                txtWatermark.FontSize = 60;
                txtWatermark.Layout = WatermarkLayout.Diagonal;
                doc.Watermark = txtWatermark;
            }

            doc.SaveToFile(output);
            //System.Diagnostics.Process.Start(output + "sample.doc");
        }

        public static void pptWaterMark(String file, String output, String pic, String text)
        {
            //加载PPT文档
            Presentation ppt = new Presentation();
            ppt.LoadFromFile(file);

            if (pic != "")
            {
                //获取图片并将其设置为页面的背景图
                Image img = Image.FromFile(pic);
                IImageData image = ppt.Images.Append(img);
                for (int i = 0; i < ppt.Slides.Count; i++)
                {
                    //为幻灯片设置背景图片类型和样式
                    ppt.Slides[i].SlideBackground.Type = Spire.Presentation.Drawing.BackgroundType.Custom;
                    ppt.Slides[i].SlideBackground.Fill.FillType = FillFormatType.Picture;
                    ppt.Slides[i].SlideBackground.Fill.PictureFill.FillType = PictureFillType.Stretch;
                    ppt.Slides[i].SlideBackground.Fill.PictureFill.Picture.EmbedImage = image;
                }
            }
            else
            {
                //获取文本的大小
                Font stringFont = new Font("Arial", 60);
                Size size = TextRenderer.MeasureText(text, stringFont);
                RectangleF rect = new RectangleF((ppt.SlideSize.Size.Width - size.Width) / 2,
                    (ppt.SlideSize.Size.Height - size.Height) / 2, size.Width, size.Height);
                for (int i = 0; i < ppt.Slides.Count; i++)
                {
                    //绘制文本，设置文本格式并将其添加到幻灯片
                    IAutoShape shape = ppt.Slides[i].Shapes.AppendShape(Spire.Presentation.ShapeType.Rectangle, rect);
                    shape.Fill.FillType = FillFormatType.None;
                    shape.ShapeStyle.LineColor.Color = Color.White;
                    shape.Rotation = -45;
                    shape.Locking.SelectionProtection = true;
                    shape.Line.FillType = FillFormatType.None;
                    shape.TextFrame.Text = text;
                    TextRange textRange = shape.TextFrame.TextRange;
                    textRange.Fill.FillType = FillFormatType.Solid;
                    textRange.Fill.SolidColor.Color = Color.FromArgb(120, Color.Gray);
                    textRange.FontHeight = 45;
                }
            }

            //保存文档
            ppt.SaveToFile(output, FileFormat.Pptx2010);
            //System.Diagnostics.Process.Start(output + "sample.pptx");
        }

        public static void pdfWaterMark(String file, String output, String pic, String text)
        {
            //加载PDF文档
            PdfDocument pdf = new PdfDocument();
            pdf.LoadFromFile(file);

            for (int i = 0; i < pdf.Pages.Count; i++)
            {
                if (pic != "")
                {
                    //获取PDF文档的第一页
                    PdfPageBase page = pdf.Pages[i];

                    //获取图片并将其设置为页面的背景图
                    Image img = Image.FromFile(pic);
                    page.BackgroundImage = img;

                    //指定背景图的位置和大小
                    page.BackgroundRegion = new RectangleF(200, 200, 200, 200);
                }
                else
                {
                    //获取PDF文档的第一页
                    PdfPageBase page = pdf.Pages[i];

                    //绘制文本，设置文本格式并将其添加到页面
                    PdfTilingBrush brush = new PdfTilingBrush(new SizeF(page.Canvas.ClientSize.Width / 2,
                        page.Canvas.ClientSize.Height / 3));
                    brush.Graphics.SetTransparency(0.3f);
                    brush.Graphics.Save();
                    brush.Graphics.TranslateTransform(brush.Size.Width / 2, brush.Size.Height / 2);
                    brush.Graphics.RotateTransform(-45);
                    PdfTrueTypeFont font = new PdfTrueTypeFont(new Font("Arial", 20f), true);
                    brush.Graphics.DrawString(text, font, PdfBrushes.Red, 0, 0,
                        new PdfStringFormat(PdfTextAlignment.Center));
                    brush.Graphics.Restore();
                    brush.Graphics.SetTransparency(1);
                    page.Canvas.DrawRectangle(brush, new RectangleF(new PointF(0, 0), page.Canvas.ClientSize));
                }
            }

            //保存文档
            pdf.SaveToFile(output);
            //System.Diagnostics.Process.Start(output + "sample.pdf");
        }
    }
}