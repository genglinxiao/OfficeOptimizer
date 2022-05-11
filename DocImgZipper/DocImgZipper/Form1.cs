using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Presentation;
using System.Drawing.Imaging;

namespace DocImgZipper
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Word Documents(*.docx; *.doc)| *.docx; *.doc|PowerPoint Presentaton(*.pptx;*.ppt)|*.pptx;*.ppt";
            openFileDialog.Multiselect = false;
            if( openFileDialog.ShowDialog()==DialogResult.OK)
            {
                textBox1.Text = openFileDialog.FileName;
                textBox1.ReadOnly = true;
                button2.Enabled = true;
            }
        }

        /// <summary>
        /// Resize the image to the specified width and height.
        /// </summary>
        /// <param name="image">The image to resize.</param>
        /// <param name="width">The width to resize to.</param>
        /// <param name="height">The height to resize to.</param>
        /// <returns>The resized image.</returns>
        public Bitmap ResizeImage(Image image, int width, int height)
        {
            var destRect = new Rectangle(0, 0, width, height);
            var destImage = new Bitmap(width, height);

            destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution);

            using (var graphics = Graphics.FromImage(destImage))
            {
                graphics.CompositingMode = CompositingMode.SourceCopy;
                graphics.CompositingQuality = CompositingQuality.HighQuality;
                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphics.SmoothingMode = SmoothingMode.HighQuality;
                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;

                using (var wrapMode = new ImageAttributes())
                {
                    wrapMode.SetWrapMode(WrapMode.TileFlipXY);
                    graphics.DrawImage(image, destRect, 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, wrapMode);
                }
            }

            return destImage;
        }

        public Stream ToStream(Image image, ImageFormat format)
        {
            var stream = new System.IO.MemoryStream();
            image.Save(stream, format);
            stream.Position = 0;
            return stream;
        }


        private void button2_Click(object sender, EventArgs e)
        {
            button2.Enabled = false;
            try
            {
                System.IO.File.Copy(textBox1.Text, textBox1.Text + ".bak");
                string filename = textBox1.Text;
                if (filename.EndsWith(".doc") || filename.EndsWith(".docx"))
                {
                    using (WordprocessingDocument document = WordprocessingDocument.Open(filename, true))
                    {

                        string tempDir = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
                        Directory.CreateDirectory(tempDir);

                        Document doc = document.MainDocumentPart.Document;

                        Dictionary<string, float> dImgZoom = new Dictionary<string, float>();

                        int iBlip = 0;

                        foreach (DocumentFormat.OpenXml.Drawing.Pictures.Picture pic in doc.Descendants<DocumentFormat.OpenXml.Drawing.Pictures.Picture>())
                        {
                            iBlip++;
                            string s = pic.OuterXml;
                            var blip = pic.BlipFill.Blip.Embed.Value;
                            ImagePart imagePart = (ImagePart)doc.MainDocumentPart.GetPartById(blip);
                            Image oriImage = Bitmap.FromStream(imagePart.GetStream());
                            long oriH = oriImage.Height;
                            long oriW = oriImage.Width;
                            float oriVRes = oriImage.VerticalResolution;
                            float oriHRes = oriImage.HorizontalResolution;

                            float spH = pic.ShapeProperties.Transform2D.Extents.Cy.Value * oriVRes / 914400;
                            float spW = pic.ShapeProperties.Transform2D.Extents.Cx.Value * oriHRes / 914400;

                            float hZoom = spH / oriH;
                            float vZoom = spW / oriW;

                            if (dImgZoom.ContainsKey(blip))
                            {
                                if (dImgZoom[blip] < Math.Max(hZoom, vZoom))
                                    dImgZoom[blip] = Math.Max(hZoom, vZoom);
                            }
                            else
                                dImgZoom[blip] = Math.Max(hZoom, vZoom);

                            //RadioButton radioButton = this.Controls.OfType<RadioButton>()
                            //    .Where(x => x.Checked).FirstOrDefault();



                            //if(radioButton!=null)
                            //{
                            //    switch(radioButton.Name)
                            //    {
                            //        case "radioButton1":

                            //            break;
                            //        case "radioButton2":
                            //            break;
                            //        case "radioButton3":
                            //            break;
                            //        case "radioButton4":
                            //            break;
                            //        case "radioButton5":
                            //            break;

                            //    }
                            //}

                            //Image newImage = ResizeImage(oriImage, 500, 500);

                            //imagePart.FeedData(ToStream(newImage, ImageFormat.Jpeg));

                            ////ImagePart newImgPart = doc.MainDocumentPart.AddImagePart(oriImage.RawFormat.ToString());


                            //string uriPath = imagePart.Uri.ToString().TrimStart('/');
                            //uriPath = uriPath.Replace('/', '\\');
                            //string tempFilename = Path.Combine(tempDir, uriPath);
                            //Directory.CreateDirectory(Path.GetDirectoryName(tempFilename));
                            //oriImage.Save(tempFilename);
                            //pic.BlipFill.Blip.Embed = doc.MainDocumentPart.GetIdOfPart(imagePart);


                        }

                        textBox3.AppendText(iBlip.ToString() + " images found in the file, referencing " + dImgZoom.Count + " actual images.");
                        textBox3.AppendText(Environment.NewLine);

                        int iImg = 0;

                        foreach (string blip in dImgZoom.Keys)
                        {
                            ImagePart imagePart = (ImagePart)doc.MainDocumentPart.GetPartById(blip);
                            Image oriImage = Bitmap.FromStream(imagePart.GetStream());

                            float tgtZoomLevel = float.Parse(textBox2.Text);

                            if (dImgZoom[blip] < tgtZoomLevel)
                            {
                                iImg++;
                                Image newImage = ResizeImage(oriImage, (int)(oriImage.Width * dImgZoom[blip] / tgtZoomLevel), (int)(oriImage.Height * dImgZoom[blip] / tgtZoomLevel));
                                imagePart.FeedData(ToStream(newImage, ImageFormat.Jpeg));
                            }
                        }
                        textBox3.AppendText(iImg.ToString() + " actual image(s) down-scaled.");
                        textBox3.AppendText(Environment.NewLine);
                    }

                }
                if (filename.EndsWith(".ppt")|| filename.EndsWith(".pptx"))
                {
                    using (PresentationDocument document = PresentationDocument.Open(filename, true))
                    {

                        string tempDir = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
                        Directory.CreateDirectory(tempDir);

                        PresentationPart doc = document.PresentationPart;

                        Dictionary<string, float> dImgZoom = new Dictionary<string, float>();

                        int iBlip = 0;

                        foreach (DocumentFormat.OpenXml.Drawing.Pictures.Picture pic in doc.Presentation.Descendants<DocumentFormat.OpenXml.Drawing.Pictures.Picture>())
                        {
                            iBlip++;
                            string s = pic.OuterXml;
                            var blip = pic.BlipFill.Blip.Embed.Value;
                            ImagePart imagePart = (ImagePart)doc.MainDocumentPart.GetPartById(blip);
                            Image oriImage = Bitmap.FromStream(imagePart.GetStream());
                            long oriH = oriImage.Height;
                            long oriW = oriImage.Width;
                            float oriVRes = oriImage.VerticalResolution;
                            float oriHRes = oriImage.HorizontalResolution;

                            float spH = pic.ShapeProperties.Transform2D.Extents.Cy.Value * oriVRes / 914400;
                            float spW = pic.ShapeProperties.Transform2D.Extents.Cx.Value * oriHRes / 914400;

                            float hZoom = spH / oriH;
                            float vZoom = spW / oriW;

                            if (dImgZoom.ContainsKey(blip))
                            {
                                if (dImgZoom[blip] < Math.Max(hZoom, vZoom))
                                    dImgZoom[blip] = Math.Max(hZoom, vZoom);
                            }
                            else
                                dImgZoom[blip] = Math.Max(hZoom, vZoom);

                            //RadioButton radioButton = this.Controls.OfType<RadioButton>()
                            //    .Where(x => x.Checked).FirstOrDefault();



                            //if(radioButton!=null)
                            //{
                            //    switch(radioButton.Name)
                            //    {
                            //        case "radioButton1":

                            //            break;
                            //        case "radioButton2":
                            //            break;
                            //        case "radioButton3":
                            //            break;
                            //        case "radioButton4":
                            //            break;
                            //        case "radioButton5":
                            //            break;

                            //    }
                            //}

                            //Image newImage = ResizeImage(oriImage, 500, 500);

                            //imagePart.FeedData(ToStream(newImage, ImageFormat.Jpeg));

                            ////ImagePart newImgPart = doc.MainDocumentPart.AddImagePart(oriImage.RawFormat.ToString());


                            //string uriPath = imagePart.Uri.ToString().TrimStart('/');
                            //uriPath = uriPath.Replace('/', '\\');
                            //string tempFilename = Path.Combine(tempDir, uriPath);
                            //Directory.CreateDirectory(Path.GetDirectoryName(tempFilename));
                            //oriImage.Save(tempFilename);
                            //pic.BlipFill.Blip.Embed = doc.MainDocumentPart.GetIdOfPart(imagePart);


                        }

                        textBox3.AppendText(iBlip.ToString() + " images found in the file, referencing " + dImgZoom.Count + " actual images.");
                        textBox3.AppendText(Environment.NewLine);

                        int iImg = 0;

                        foreach (string blip in dImgZoom.Keys)
                        {
                            ImagePart imagePart = (ImagePart)doc.MainDocumentPart.GetPartById(blip);
                            Image oriImage = Bitmap.FromStream(imagePart.GetStream());

                            float tgtZoomLevel = float.Parse(textBox2.Text);

                            if (dImgZoom[blip] < tgtZoomLevel)
                            {
                                iImg++;
                                Image newImage = ResizeImage(oriImage, (int)(oriImage.Width * dImgZoom[blip] / tgtZoomLevel), (int)(oriImage.Height * dImgZoom[blip] / tgtZoomLevel));
                                imagePart.FeedData(ToStream(newImage, ImageFormat.Jpeg));
                            }
                        }
                        textBox3.AppendText(iImg.ToString() + " actual image(s) down-scaled.");
                        textBox3.AppendText(Environment.NewLine);
                    }

                }


            }
            catch (Exception exp)
            {
                textBox3.AppendText(exp.Message);
                textBox3.AppendText(Environment.NewLine);
            }
            textBox3.AppendText("Process done!");
            textBox3.AppendText(Environment.NewLine);
            button2.Enabled = true;
        }
    }
}
