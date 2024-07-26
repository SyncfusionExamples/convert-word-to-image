using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;

namespace First_page_of_Word_to_image
{
    class Program
    {
        static void Main(string[] args)
        {
			//Open the file as stream
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open))
            {
                //Load an existing Word document.
                using (WordDocument wordDocument = new WordDocument(fileStream, FormatType.Automatic))
                {
                    //Create an instance of DocIORenderer.
                    using (DocIORenderer renderer = new DocIORenderer())
                    {
                        //Convert the first page of Word document into an image.
                        Stream imageStream = wordDocument.RenderAsImages(0, ExportImageFormat.Jpeg);                        
                        //Create the output image file stream.
                        using (FileStream fileStreamOutput = File.Create(Path.GetFullPath(@"../../../WordToImage.jpeg")))
                        {
                            //Copy the converted image stream into created output stream.
                            imageStream.CopyTo(fileStreamOutput);
                        }
                    }
                }
            }
        }
    }
}
