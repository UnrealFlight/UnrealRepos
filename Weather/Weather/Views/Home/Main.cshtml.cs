using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Identity.Client;
using Weather.Models;

namespace Weather.Views.Home
{
    public class MainModel : PageModel
    {
        
        public class WeatherList
        {
            //public int id { get; set; }
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
        //public List<WeatherList> weatherLists { get; set; }
        public void OnGet()
        {
         

        }
    }
}
