using DocumentFormat.OpenXml;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;
using System.Linq;
namespace openXmlDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            const string templatePath = @"C:\Users\Gaurav Koli\source\repos\openXmlDemo\openXmlDemo\target\SampleDocOriginal.docx";
            const string resultPath = @"C:\Users\Gaurav Koli\source\repos\openXmlDemo\openXmlDemo\Result\result.docx";
            try
            {

                using (WordprocessingDocument wordDocument = WordprocessingDocument.CreateFromTemplate(templatePath, true))
                {
                    ReplaceParagraphParts(wordDocument.MainDocumentPart.Document, wordDocument);

                    wordDocument.SaveAs(resultPath);
                }
            }
            catch (IOException ioe)
            {
                Console.WriteLine(ioe.Message);
                Console.ReadKey();
            }
        }
       
        private static void ReplaceParagraphParts(OpenXmlElement element, WordprocessingDocument wordDocument)
        {
            //int i = 1;
         
            // Getting all Paragraph in Xml File

            Drawing draw = element.Descendants<Drawing>().FirstOrDefault();
            FileInfo fileInfo = new FileInfo("C:\\Users\\Gaurav Koli\\Downloads\\battlefield_bad_company_2_table_room_parquet-740403.jpg");
            string embed = null;
            DocumentFormat.OpenXml.Drawing.Blip blip=null;

            foreach (var paragraph in element.Descendants<Paragraph>())
            {
                //Getting blip Id to get Image Part
                SdtAlias sa = paragraph.Descendants<SdtAlias>().SingleOrDefault();
                if (sa != null && sa.Val == "crmndc_signatureurl")
                {
                    
                    sa.Val = "Change Picture";
                    Console.WriteLine("Done");
                    Drawing dr = paragraph.Descendants<Drawing>().FirstOrDefault();
                    //9525 is EMU per pixel
                    Int64 finalCx= (600 * 9525);
                    Int64 finalCy = (300 * 9525);

                    //resize the image
                    dr.Inline.Extent.Cx = finalCx;
                    dr.Inline.Extent.Cy= finalCy;

                    dr.Inline.Graphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>().ShapeProperties.Transform2D.Extents.Cx = finalCx;
                    dr.Inline.Graphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>().ShapeProperties.Transform2D.Extents.Cy = finalCy;

                    if (dr != null)
                    {
                        blip = dr.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
                        if (blip != null)
                            embed = blip.Embed;
                    }
                }

                //Getting Image part and change the Image
                
                if (embed != null)
                {
                    IdPartPair idpp = wordDocument.MainDocumentPart.Parts.Where(pa => pa.RelationshipId == embed).FirstOrDefault();
                    if (idpp != null)
                    {
                        ImagePart ip = (ImagePart)idpp.OpenXmlPart;
                        try
                        {
                            using (FileStream fileStream = fileInfo.OpenRead())
                            {
                                
                                ip.FeedData(fileStream);
                              // fileStream.Close();

                            }
                            if (blip != null)
                                blip.Embed.Value = wordDocument.MainDocumentPart.GetIdOfPart(ip);
                            Console.WriteLine("done " + wordDocument.MainDocumentPart.GetIdOfPart(ip));
                           // Console.ReadKey();
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.StackTrace);
                            Console.ReadKey();
                        }

                    }
                    embed = null;
                }

                //Changing the Templete Text
                var sdtContentText = paragraph.Descendants<Text>();
               
                if (sdtContentText != null) {
                    foreach (Text text in sdtContentText)
                    {
                        switch (text.Text)
                        {
                            case "<<crmndc_seller1_fullname>>":
                                text.Text = "Common" + i;
                                i++;
                                break;

                            case "Lysaker, ":
                                text.Text = "Common" + i;
                                i++;
                                break;

                            case "21.03.2016":
                                text.Text = "Common" + i;
                                i++;
                                break;

                            case "<<title>>":
                                text.Text = "Common" + i;
                                i++;
                                break;

                            case "<<crmndc_buyer1_fullname>>":
                                text.Text = "Common" + i;
                                i++;
                                break;

                            case "crmndc_insurancecompany_name":
                                text.Text = "Common" + i;
                                i++;
                                break;

                            case "<<":
                                text.Text = "";
                                break;

                            case ">>":
                                text.Text = "";
                                break;

                            default:
                                break;

                        }
                       
                    }
                }
               
               
 
            }

            
        }


       
    }
}
    

