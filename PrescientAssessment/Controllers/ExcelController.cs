using Microsoft.AspNetCore.Mvc;
using PrescientAssessment.Service;

namespace PrescientAssessment.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        IExcelService _service;
        public ExcelController(IExcelService service)
        {
            _service = service;
        }

        [HttpGet("downloadFiles")]
        public IActionResult DownloadFiles()
        {
            _service.GetDownloadedList();
            return Ok("Finished downloading files");
        }

        [HttpGet("readFiles")]
        public IActionResult ReadFiles()
        {
            _service.ReadExcelFiles();
            return Ok("Finished reading files");
        }
    }
}
