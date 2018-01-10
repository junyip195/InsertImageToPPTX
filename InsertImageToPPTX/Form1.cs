using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
//using DocumentFormat.OpenXml.Presentation;
using Spire.Presentation;
using Spire.Presentation.Drawing;
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
        List<Image> images = null;
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
                            images.Add(Image.FromStream(ms));
                        }

                    }

                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Presentation presentation = new Presentation();
            presentation.LoadFromFile(fileName);
            IMasterSlide master = presentation.Masters[0];
            String image = @"logo.png";
            RectangleF rff = new RectangleF(40, 40, 100, 80);
            IEmbedImage pic = master.Shapes.AppendEmbedImage(ShapeType.Rectangle, image, rff);
            pic.Line.FillFormat.FillType = FillFormatType.None;
            presentation.Slides.Append();
            presentation.SaveToFile("result.pptx", FileFormat.Auto);
            System.Diagnostics.Process.Start("result.pptx");
        }

    }

}

