//---------------
//Это файл с попыткой вникнуть в НПОИ, но чет так и не получилось... Не дружит с FileStream...
//---------------





//using Microsoft.AspNetCore.Mvc;
//using System.Diagnostics;
//using Weather.Models;
//using Microsoft.EntityFrameworkCore;
//using System.IO;
//using System.Data.SqlClient;
//using System.Web;
//using Microsoft.Data.SqlClient.DataClassification;
//using Microsoft.AspNetCore.Http;
//using Microsoft.AspNetCore.Hosting;
//using Microsoft.Extensions.Logging;
//using Microsoft.Data.SqlClient;
//using NPOI.SS.UserModel;
//using NPOI.XSSF.UserModel;
//using System.Data;
//using NPOI.HSSF.Record.Chart;

//namespace Weather.Controllers
//{
//    public class HomeController : Controller
//    {

//        //ApplicationContext _context;
//        IWebHostEnvironment _appEnvironment;
//        //private readonly ILogger<HomeController> _logger;

//        //ApplicationContext db;
//        public HomeController(IWebHostEnvironment appEnvironment)
//        {

//            _appEnvironment = appEnvironment;
//        }

//        //public HomeController(ILogger<HomeController> logger)
//        //{
//        //    _logger = logger;
//        //}

//        public IActionResult Index()
//        {
//            return View();
//        }
//        public IActionResult Privacy(/*ApplicationContext context*/)
//        {
//            //_context = context;
//            return View(/*_context.Files.ToList()*/);
//        }
//        [HttpPost]
//        public IActionResult Privacy(IFormFile FileUpload)
//        {

//            // List<WeatherList> weatherLists= new List<WeatherList>();
//            WeatherList weatherLists = new WeatherList();
//            XSSFWorkbook? xssxfwb = null;
//            string path = "\\Files\\" + FileUpload.FileName;

//            //using (var file = new FileStream(_appEnvironment.WebRootPath + path, FileMode.Create))
//            //{
//            //    FileUpload.CopyToAsync(file);
//            //}            
//            using (var file = new FileStream("D:\\" + FileUpload.FileName, FileMode.Create))
//            {
//                FileUpload.CopyToAsync(file);
//            }
//            // string s = "D:\\" + FileUpload.FileName.ToString();
//            //string s = "D:\\moskva_2012.xlsx";
//            string s = "D:\\moskva_2012.xlsx";
//            //using (var filestream = new FileStream("D:"+@"\"+"moskva_2012.xlsx", FileMode.Open, FileAccess.Read))
//            //{
//            //    filestream.Position = 0;
//            //    xssxfwb = new XSSFWorkbook(filestream);
//            //}
//            FileStream fileStream = new FileStream("D:\\moskva_2012.xlsx", FileMode.Open);

//            fileStream.Position = 0;
//            xssxfwb = new XSSFWorkbook(fileStream);
//            fileStream.Close();

//            ISheet sheet = xssxfwb.GetSheetAt(0);
//            //int lastcell = row.LastCellNum;
//            //12

//            SqlConnection sqlcon = new SqlConnection();
//            sqlcon = new SqlConnection(@"Data Source=DESKTOP-CP0UKNL\SQLEXPRESS;database=Weather;User Id=sa;Password=1234;Encrypt=yes;TrustServerCertificate=true;");
//            int id = 0;
//            bool ex = false;
//            SqlCommand comand = new SqlCommand();
//            sqlcon.Open();
//            for (int j = 0; j < 12; j++)
//            {
//                sheet = xssxfwb.GetSheetAt(j);
//                for (int i = 4; i <= sheet.LastRowNum; i++)
//                {
//                    ex = false;
//                    IRow row = sheet.GetRow(i);
//                    weatherLists.date = row.GetCell(0).StringCellValue;
//                    weatherLists.time = row.GetCell(1).ToString();
//                    weatherLists.T = row.GetCell(2).ToString();
//                    weatherLists.humidity = row.GetCell(3).ToString();
//                    weatherLists.Td = row.GetCell(4).ToString();
//                    weatherLists.pressure = row.GetCell(5).ToString();
//                    weatherLists.direction = row.GetCell(6).ToString();
//                    weatherLists.speed = row.GetCell(7).ToString();
//                    weatherLists.sky = row.GetCell(8).ToString();
//                    weatherLists.h = row.GetCell(9).ToString();
//                    weatherLists.VV = row.GetCell(10).ToString();
//                    try
//                    {
//                        weatherLists.other = row.GetCell(11).ToString();
//                    }
//                    catch { weatherLists.other = ""; }

//                    while (ex != true)
//                    {
//                        string sqladd = "INSERT INTO Weather (id,date,time,T,humidity,Td,pressure,direction,speed,sky,h,VV,other) VALUES ('" + id + "','" + weatherLists.date + "','" + weatherLists.time + "','" + weatherLists.T + "','" + weatherLists.humidity + "','" + weatherLists.Td + "','" + weatherLists.pressure + "','" + weatherLists.direction + "','" + weatherLists.speed + "','" + weatherLists.sky + "','" + weatherLists.h + "','" + weatherLists.VV + "','" + weatherLists.other + "')";

//                        comand.CommandText = sqladd;

//                        comand.Connection = sqlcon;
//                        try
//                        {
//                            comand.ExecuteNonQuery();
//                            ex = true;
//                        }
//                        catch { id += 1; }
//                    }
//                }
//            }

//            sqlcon.Close();
//            return View();
//        }
//        //public async Task<IActionResult> addFile( IFormFile FileUpload)
//        //{

//        //    if (FileUpload != null)
//        //    {
//        //        // путь к папке Files
//        //        string path = "/Files/" + FileUpload.FileName;
//        //        // сохраняем файл в папку Files в каталоге wwwroot
//        //        using (var fileStream = new FileStream(_appEnvironment.WebRootPath + path, FileMode.Create))
//        //        {
//        //            await FileUpload.CopyToAsync(fileStream);
//        //        }
//        //        fileModel file = new fileModel { name = FileUpload.FileName, path = path };
//        //        _context.Files.Add(file);
//        //        _context.SaveChanges();
//        //    }

//        //    return RedirectToAction("Index");
//        //}
//        public IActionResult Main()
//        {
//            return View();
//        }

//        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
//        public IActionResult Error()
//        {
//            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
//        }
//    }
//}