using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ExelParser.Services;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace ExelParser.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        ExcelService _excelService;

        public ExcelController(ExcelService excelService)
        {
            _excelService = excelService;
        }

        // GET api/<ExcelController>/5
        [HttpGet("{name}")]
        public IActionResult Get(string name)
        {
            var dataTable = _excelService.ReadExcel(name);

            var jsonDurty = JsonConvert.SerializeObject(dataTable, Formatting.Indented);

            return Content(jsonDurty.ToString(), contentType:"application/json");
        }
    }
}
