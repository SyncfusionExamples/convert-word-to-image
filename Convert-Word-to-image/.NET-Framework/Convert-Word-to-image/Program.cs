using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;
using Syncfusion.OfficeChartToImageConverter;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;

namespace Convert_Word_to_image
{
    class Program
    {
        static void Main(string[] args)
        {
            //Load an existing Word document.
            using (WordDocument wordDocument = new WordDocument(Path.GetFullPath(@"../../Template.docx"), FormatType.Docx))
            {
                //Initialize the ChartToImageConverter for converting charts during Word to image conversion.
                wordDocument.ChartToImageConverter = new ChartToImageConverter();
               //Convert an entire Word document to images.
                Image[] images = wordDocument.RenderAsImages(ImageType.Bitmap);
                for (int i=0; i<images.Length; i++)
                {
                    //Save the images as jpeg.
                    images[i].Save(Path.GetFullPath(@"../../WordToImage_" + i + ".jpeg"), ImageFormat.Jpeg);
                }
            }
        }
    }
}
