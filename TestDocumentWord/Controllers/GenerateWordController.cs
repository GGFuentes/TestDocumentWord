using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Net;
using TestDocumentWord.Repositories;

namespace TestDocumentWord.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class GenerateWordController : ControllerBase
    {
        private readonly IGenerateWordRepository _generateWordRepository;
        public GenerateWordController(IGenerateWordRepository generateWordRepository)
        {
            _generateWordRepository = generateWordRepository;
        }

        [HttpPost]
        public async Task<IActionResult> GenerateWord(string filePath)
        {
            return Ok(await _generateWordRepository.GenerateDocument(filePath));
           
        }

    }
}
