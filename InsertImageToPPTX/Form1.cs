using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace InsertImageToPPTX
{
    public partial class Form1 : Form
    {
        string fileName = null;
        List<string> images = null;
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog open_file_dialog = new OpenFileDialog();

            if (open_file_dialog.ShowDialog() == DialogResult.OK)
            {
                // Get full & safe filename.
                string full_filename = open_file_dialog.FileName;
                string safe_filename = open_file_dialog.SafeFileName;
                textBox1.Text = fileName;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string templateUrl = @"https://www.google.co.uk/search?q={0}&tbm=isch&site=imghp";

            //check that we have a term to search for.
            if (string.IsNullOrEmpty(textBox1.Text))
            {
                MessageBox.Show("Please supply a search term"); return;
            }
            else
            {
                using (WebClient wc = new WebClient())
                {
                    //lets pretend we are IE8 on Vista.
                    wc.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0)");
                    string result = wc.DownloadString(String.Format(templateUrl, new object[] { textBox1.Text }));

                    //we have valid markup, this will change from time to time as google updates.
                    if (result.Contains("images_table"))
                    {
                        HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                        doc.LoadHtml(result);

                        //lets create a linq query to find all the img's stored in that images_table class.
                        /*
                         * Essentially we get search for the table called images_table, and then get all images that have a valid src containing images?
                         * which is the string used by google
                        eg  https://encrypted-tbn3.gstatic.com/images?q=tbn:ANd9GcQmGxh15UUyzV_HGuGZXUxxnnc6LuqLMgHR9ssUu1uRwy0Oab9OeK1wCw
                         */

                        var imgList = from tables in doc.DocumentNode.Descendants("table")
                                      from img in tables.Descendants("img")
                                      where tables.Attributes["class"] != null && tables.Attributes["class"].Value == "images_table"
                                      && img.Attributes["src"] != null && img.Attributes["src"].Value.Contains("images?")
                                      select img;



                        byte[] downloadedData = wc.DownloadData(imgList.First().Attributes["src"].Value);

                        if (downloadedData != null)
                        {
                            //store the downloaded data in to a stream
                            System.IO.MemoryStream ms = new System.IO.MemoryStream(downloadedData, 0, downloadedData.Length);

                            //write to that stream the byte array
                            ms.Write(downloadedData, 0, downloadedData.Length);

                            //load an image from that stream.
                            pictureBox1.Image = Image.FromStream(ms);
                        }

                    }

                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {

        }
        public class SegmentationShapeProperties
        {
            public Int64 OffsetX { get; set; }
            public Int64 OffsetY { get; set; }
            public Int64 ScaleX { get; set; }
            public Int64 ScaleY { get; set; }
        }
        public class SegmentationSlideInputData
        {
            public int SlideId { get; set; }
            public OpenXmlUtils.SegmentationShapeProperties ShapeProperties { get; set; }
        }

        public void InsertImages(List<string> imageFilesWithPath, string presentation)
        {
            using (PresentationDocument prstDoc = PresentationDocument.Open(presentation, true))
            {
                PresentationPart presentationPart = prstDoc.PresentationPart;
                var slideParts = OpenXmlUtils.GetSlidePartsInOrder(presentationPart); //gets all the slide parts present in the documetn
                Slide slide = null;
                foreach (string imageWithPath in imageFilesWithPath)
                {
                    SegmentationSlideInputData data = GetWorkingImageDetails(Path.GetFileName(imageWithPath));//function which decides which slide to work on and image scaling options
                    slide = slideParts.ElementAt(data.SlideId).Slide;
                    Picture pic = OpenXmlUtils.AddPicture(slide, imageWithPath, data.ShapeProperties);
                    slide.Save();
                }
                prstDoc.PresentationPart.Presentation.Save();
            }
        }

        private SegmentationSlideInputData GetWorkingImageDetails(string fileName)
        {
            SegmentationSlideInputData data = new SegmentationSlideInputData();
            data.SlideId = 0;//slide id to work on
            data.ShapeProperties = new OpenXmlUtils.SegmentationShapeProperties() { OffsetX = 4695825L, OffsetY = 504825L, ScaleX = 6721828L, ScaleY = 1988495L };//offset specifies the position, scale specifies the height and widht of image
            break;

            return data;
        }


        internal static P.Picture AddPicture(this Slide slide, string imageFile, SegmentationShapeProperties sP)
        {
            P.Picture picture = new P.Picture();

            string embedId = string.Empty;
            UInt32Value picId = 10001U;
            string name = string.Empty;

            if (slide.Elements<P.Picture>().Count() > 0)
            {
                picId = ++slide.Elements<P.Picture>().ToList().Last().NonVisualPictureProperties.NonVisualDrawingProperties.Id;
            }
            name = "image" + picId.ToString();
            embedId = "rId" + (RandomString(5)).ToString(); // some value

            P.NonVisualPictureProperties nonVisualPictureProperties = new P.NonVisualPictureProperties()
            {
                NonVisualDrawingProperties = new P.NonVisualDrawingProperties() { Name = name, Id = picId, Title = name },
                NonVisualPictureDrawingProperties = new P.NonVisualPictureDrawingProperties() { PictureLocks = new Z.Drawing.PictureLocks() { NoChangeAspect = true } },
                ApplicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties() { UserDrawn = true }
            };

            P.BlipFill blipFill = new P.BlipFill() { Blip = new Z.Drawing.Blip() { Embed = embedId } };
            Z.Drawing.Stretch stretch = new Z.Drawing.Stretch() { FillRectangle = new Z.Drawing.FillRectangle() };
            blipFill.Append(stretch);

            P.ShapeProperties shapeProperties = new P.ShapeProperties()
            {
                Transform2D = new Z.Drawing.Transform2D()
                {
                    //Offset = new Z.Drawing.Offset() { X = 1835696L, Y = 1036712L },
                    //Extents = new Z.Drawing.Extents() { Cx = 5334617, Cy = 1025963 }
                    Offset = new Z.Drawing.Offset() { X = sP.OffsetX, Y = sP.OffsetY },
                    Extents = new Z.Drawing.Extents() { Cx = sP.ScaleX, Cy = sP.ScaleY }
                }
            };
            Z.Drawing.PresetGeometry presetGeometry = new Z.Drawing.PresetGeometry() { Preset = Z.Drawing.ShapeTypeValues.Rectangle };
            Z.Drawing.AdjustValueList adjustValueList = new Z.Drawing.AdjustValueList();

            presetGeometry.Append(adjustValueList);
            shapeProperties.Append(presetGeometry);
            picture.Append(nonVisualPictureProperties);
            picture.Append(blipFill);
            picture.Append(shapeProperties);

            slide.CommonSlideData.ShapeTree.Append(picture);

            // Add Image part
            slide.AddImagePart(embedId, imageFile);

            slide.Save();
            return picture;
        }


        private static void AddImagePart(this Slide slide, string relationshipId, string imageFile)
        {
            ImagePart imgPart = slide.SlidePart.AddImagePart(GetImagePartType(imageFile), relationshipId);
            using (FileStream imgStream = File.Open(imageFile, FileMode.Open))
            {
                imgPart.FeedData(imgStream);
            }
        }
    }
}
}

