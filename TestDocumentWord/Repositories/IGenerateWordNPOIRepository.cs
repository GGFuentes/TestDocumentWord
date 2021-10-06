using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using TestDocumentWord.CrossCutting;

namespace TestDocumentWord.Repositories
{
    public interface IGenerateWordNPOIRepository
    {
        Task<Response<string>> GenerateDocument(string filePath);
    }
}
