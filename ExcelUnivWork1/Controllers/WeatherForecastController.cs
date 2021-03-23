using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;

namespace ExcelUnivWork1.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class WeatherForecastController : ControllerBase
    {
        private readonly ILogger<WeatherForecastController> _logger;
        private ExcelService _excelService;
        
        public WeatherForecastController(ILogger<WeatherForecastController> logger, ExcelService excelService)
        {
            _logger = logger;
            _excelService = excelService;
        }

        [HttpPut()]
        [DisableRequestSizeLimit]
        public async Task<IActionResult> GetExcelTemplateForCarrier(
            Request request)
        {
            ActionResult result = null;

            try
            {
                result = new FileStreamResult(await _excelService.Get(request),
                             "application/vnd.ms-excel")
                         {
                             FileDownloadName = "dataset.xlsx",
                         };

                Response.StatusCode = (int)HttpStatusCode.OK;
            }
            catch (Exception exception)
            {
                Response.StatusCode = (int)HttpStatusCode.BadRequest;
                var message = "Failed to convert";

                if (!string.IsNullOrWhiteSpace(exception.Message))
                {
                    message = exception.Message;
                }
                
            }

            return result;
        }
    }
}