using System.IO;

namespace GenerateCalendar.Services
{
    public interface ICalendarService
    {
        MemoryStream GeneratedPackage();
    }
}