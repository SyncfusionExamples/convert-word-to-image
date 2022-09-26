using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;

namespace Convert_Word_to_image
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open the file as Stream. 
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open))
            {
                //Load an existing Word document.
                using (WordDocument wordDocument = new WordDocument(fileStream, FormatType.Automatic))
                {
                    //Create a new instance of DocIORenderer. 
                    using (DocIORenderer renderer = new DocIORenderer())
                    {
                        //Convert an entire Word document to images.
                        Stream[] imageStreams = wordDocument.RenderAsImages();
                        for (int i=0; i< imageStreams.Length;i++)
                        {
                            //Create the output image file stream.
                            using (FileStream fileStreamOutput = File.Create(Path.GetFullPath(@"../../../WordToImage_" + i + ".jpeg")))
                            {
                                //Copy the converted image stream into created output stream.
                                imageStreams[i].CopyTo(fileStreamOutput);
                            }
                        }
                    }
                }
            }
        }
    }
}
