using Microsoft.EntityFrameworkCore;

namespace Weather.Models
{
    
    public class WeatherList
    {
        public string? date { get; set; }
        public string? time { get; set; }
        public string? T { get; set; }
        public string? humidity { get; set; }
        public string? Td { get; set; }
        public string? pressure { get; set; }
        public string? direction { get; set; }
        public string? speed { get; set; }
        public string? sky { get; set; }
        public string? h { get; set; }
        public string? VV { get; set; }
        public string? other { get; set; }
    }
    public class fileModel
    {
        public int id { get; set; }
        public string name { get; set; }
        public string path { get; set; }
    }
   
}