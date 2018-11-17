using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using GenerateCalendar.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace GenerateCalendar.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class CalendarController : ControllerBase
    {
        private readonly ICalendarService _calendarService;

        public CalendarController(ICalendarService calendarService)
        {
            _calendarService = calendarService;
        }


        [HttpGet]
        public IActionResult GetCalendar()
        {
            var ms = _calendarService.GeneratedPackage();
            var filename = "calendar.docx";
            var fileContentResult = new FileContentResult(ms.ToArray(), "application/octet-stream")
            {
                FileDownloadName = filename
            };

            return fileContentResult;
        }

    }
}