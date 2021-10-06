using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Hosting;
using System;
using System.IO;
using System.Threading.Tasks;
using TestDocumentWord.CrossCutting;

namespace TestDocumentWord.Repositories
{
    public class GenerateWordRepository : IGenerateWordRepository
    {
        private readonly IWebHostEnvironment ruta;
        public GenerateWordRepository(IWebHostEnvironment _ruta)
        {
            ruta = _ruta;
        }

        public async Task<Response<string>> GenerateDocument(string filepath)
        {
            var response = new Response<string>();
            try
            {
                filepath = $"{this.ruta.WebRootPath}/ArchivosExcel/{DateTime.Now.ToString("yyyyMMddHHmm")}.docx";
                // Create a document by supplying the filepath. 
                using (WordprocessingDocument wordDocument =
                    WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document))
                {
                    // Add a main document part. 
                    MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                    // Create the document structure and add some text.
                    mainPart.Document = new Document();
                    Body body = mainPart.Document.AppendChild(new Body());
                    Paragraph para = body.AppendChild(new Paragraph());
                    Run run = para.AppendChild(new Run());
                    run.AppendChild(new Text("Create text in body - CreateWordprocessingDocument"));
                    response.Message= "ok";
                    response.isSuccess= true;
                    response.isError= false;
                    
                }
                byte[] byteArray = File.ReadAllBytes(filepath);
                string base64 = Convert.ToBase64String(byteArray);
                //DeleteFileCreated 
                //File.Delete(filepath);
                response.Data = base64;
            }
            catch(Exception e)
            {
                response.Message = "ok";
                response.isSuccess = true;
                response.isError = false;
                response.Data = "";
            }
            return response;
        }
    }
}
