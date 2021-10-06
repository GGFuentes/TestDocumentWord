using Microsoft.AspNetCore.Mvc;
using System.Threading.Tasks;
using TestDocumentWord.Repositories;

namespace TestDocumentWord.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class GenerateWordNPOIController : ControllerBase
    {
        private readonly IGenerateWordNPOIRepository _generateWordNPOIRepository;
        public GenerateWordNPOIController(IGenerateWordNPOIRepository generateWordNPOIRepository)
        {
            _generateWordNPOIRepository = generateWordNPOIRepository;
        }

        [HttpPost]
        public async Task<IActionResult> GenerateWord(string filePath)
        {
            return Ok(await _generateWordNPOIRepository.GenerateDocument(filePath));
           
        }

    }
}
