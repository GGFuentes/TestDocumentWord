using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Hosting;
using NPOI.XWPF.UserModel;
using System;
using System.IO;
using System.Threading.Tasks;
using TestDocumentWord.CrossCutting;

namespace TestDocumentWord.Repositories
{
    public class GenerateWordNPOIRepository : IGenerateWordNPOIRepository
    {
        private readonly IWebHostEnvironment ruta;
        public GenerateWordNPOIRepository(IWebHostEnvironment _ruta)
        {
            ruta = _ruta;
        }

        public async Task<Response<string>> GenerateDocument(string filepath)
        {
            var response = new Response<string>();
            try
            {
                filepath = $"{this.ruta.WebRootPath}/ArchivosExcel/{DateTime.Now.ToString("yyyyMMddHHmm")}.docx";

                XWPFDocument doc = new XWPFDocument();
                XWPFParagraph p2 = doc.CreateParagraph();
                XWPFRun r2 = p2.CreateRun();
                r2.SetText("test");
                r2.SetText("test");


                var widthEmus = (int)(400.0 * 9525);
                var heightEmus = (int)(300.0 * 9525);

                using (FileStream picData = new FileStream($"{this.ruta.WebRootPath}/Image/HumpbackWhale.jpg", FileMode.Open, FileAccess.Read))
                {
                    r2.AddPicture(picData, (int)PictureType.PNG, "image1", widthEmus, heightEmus);
                }

                FileStream sw = File.Create(filepath);
                
                    doc.Write(sw);     
               
                // Create a document by supplying the filepath.               
                byte[] byteArray = File.ReadAllBytes(filepath);
                string base64 = Convert.ToBase64String(byteArray);
                //DeleteFileCreated 
               // File.Delete(filepath);
                response.Data = base64;
            }
            catch(Exception e)
            {
                response.Message = e.Message;
                response.isSuccess = true;
                response.isError = false;
                response.Data = "";
            }
            return response;
        }
    }
}
