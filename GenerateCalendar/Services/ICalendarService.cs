using System.IO;

namespace GenerateCalendar.Services
{
    public interface ICalendarService
    {
        MemoryStream GeneratePackage(int year);
    }
}