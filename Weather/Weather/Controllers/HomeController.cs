using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using Weather.Models;
using Microsoft.EntityFrameworkCore;
using System.IO;
using System.Data.SqlClient;
using System.Web;
using Microsoft.Data.SqlClient.DataClassification;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Data.SqlClient;
using System.Data;
using System.Collections.Generic;
using Aspose.Cells;
using Microsoft.VisualBasic;
using Microsoft.AspNetCore.Mvc.Rendering;

namespace Weather.Controllers
{
    public class HomeController : Controller
    {

        List<WeatherList> weatherLists = new List<WeatherList>();
        List<WeatherList> weatherLists2 = new List<WeatherList>();
        Workbook wb;
        //для получения пути к wwwroot
        IWebHostEnvironment _appEnvironment;

        public HomeController(IWebHostEnvironment appEnvironment)
        {
            _appEnvironment = appEnvironment;
        }

        [HttpGet]
        public IActionResult Main()
        {
            //при отрытии html main добавить нулевой элемент в лист
            WeatherList weather = new WeatherList();
            weather.date = "";
            weather.time = "";
            weather.T = "";
            weather.humidity = "";
            weather.Td = "";
            weather.pressure = "";
            weather.direction = "";
            weather.speed = "";
            weather.sky = "";
            weather.h = "";
            weather.VV = "";
            weather.other = "";
            weatherLists.Add(weather);
            string strr = weatherLists[0].date;
            ViewBag.People = new List<WeatherList>(weatherLists);
            return View();
        }
        public IActionResult Index()
        {
            return View();
        }
        public IActionResult Privacy()
        {
            return View();
        }

        [HttpPost]
        public IActionResult Privacy(IFormFile FileUpload)
        {
            string str;
            //filename передает тип file.xlsx
            string[] filename= FileUpload.FileName.Split('.');
            bool del = false;

            if (FileUpload == null)
            {
                return View();
            }

            //создание листа и добавление в него только имен файлов из каталога
            List<string> files =new List<string>(  Directory.GetFiles(_appEnvironment.WebRootPath+"\\Files\\").Select(x=>Path.GetFileNameWithoutExtension(x)));
           
            //проверка на уже сущеструющий файл
            for (int i=0;i<files.Count; i++)
            {
                str = files[i];
                if (str == filename[0])
                {
                    return View();
                }
            }
            
            WorksheetCollection wsc;
            Worksheet ws;
            string path = "\\Files\\" + FileUpload.FileName;

            //скопировать(загрузить на сервер) файл
            using (var file = new FileStream(_appEnvironment.WebRootPath+path, FileMode.Create))
            {
                FileUpload.CopyTo(file);
            }

            //открыть скопированную книу
            string s = _appEnvironment.WebRootPath+path;
            wb = new Workbook(s);
            wsc = wb.Worksheets;
            ws = wsc[0];

            int columns = ws.Cells.MaxDataColumn;
            int rows = ws.Cells.MaxDataRow;

            if (columns != 11 && rows <5)
            {
                return View();
            }

            SqlConnection sqlcon = new SqlConnection();
            sqlcon = new SqlConnection(@"Data Source=DESKTOP-CP0UKNL\SQLEXPRESS;database=Weather;User Id=sa;Password=1234;Encrypt=yes;TrustServerCertificate=true;");
            SqlCommand comand = new SqlCommand();

            //id для новых строк sql
            int id = 0;

            //bool переменная для проверки уже существующего id
            bool ex = false;


            sqlcon.Open();

            //Запись данных в sql
            for (int j = 0; j <= columns; j++)
            {
                    ws = wsc[j];
               
                for (int i = 4; i < rows; i++)
                {
                    ex = false;

                    WeatherList weatherLists = new WeatherList();

                    weatherLists.date = ws.Cells[i, 0].StringValue;
                    weatherLists.time = ws.Cells[i, 1].StringValue;
                    weatherLists.T = ws.Cells[i, 2].StringValue;
                    weatherLists.humidity = ws.Cells[i, 3].StringValue;
                    weatherLists.Td = ws.Cells[i, 4].StringValue;
                    weatherLists.pressure = ws.Cells[i, 5].StringValue;
                    weatherLists.direction = ws.Cells[i, 6].StringValue;
                    weatherLists.speed = ws.Cells[i, 7].StringValue;
                    weatherLists.sky = ws.Cells[i, 8].StringValue;
                    weatherLists.h = ws.Cells[i, 9].StringValue;
                    weatherLists.VV = ws.Cells[i, 10].StringValue;
                    try
                    {
                        weatherLists.other = ws.Cells[i, 11].Value.ToString();
                    }
                    catch { weatherLists.other = ""; }

                    while (ex != true)
                    {
                        string sqladd = "INSERT INTO Weather (id,date,time,T,humidity,Td,pressure,direction,speed,sky,h,VV,other) VALUES ('" + id + "','" + weatherLists.date + "','" + weatherLists.time + "','" + weatherLists.T + "','" + weatherLists.humidity + "','" + weatherLists.Td + "','" + weatherLists.pressure + "','" + weatherLists.direction + "','" + weatherLists.speed + "','" + weatherLists.sky + "','" + weatherLists.h + "','" + weatherLists.VV + "','" + weatherLists.other + "')";

                        comand.CommandText = sqladd;

                        comand.Connection = sqlcon;
                        try
                        {
                            comand.ExecuteNonQuery();
                            ex = true;
                        }
                        catch { id += 1; }
                    }
                }
            }

            sqlcon.Close();
            return View();
        }

        [HttpPost]
        public IActionResult Main(string calendar)
        {
            string chosedate = calendar;
            //2014-01

            //Проверка введения даты
            if (calendar != null)
            {
                string[] words = chosedate.Split('-');
                string str = words[0];
                string str2 = words[1];
                chosedate = str2 + "." + str; //месяц и год
            }

            string con = "Data Source=DESKTOP-CP0UKNL\\SQLEXPRESS;database=weather;User Id=sa;Password = 1234;Encrypt=yes;TrustServerCertificate=true;";
            string sqlselect = "";
            string sqlselect1 = "SELECT * FROM Weather";
            string sqlselect2 = "SELECT * FROM Weather WHERE date LIKE '%"+chosedate+"'";

            if (calendar != null)
            {
                sqlselect = sqlselect2;
            }
            else
            {
                sqlselect = sqlselect1;
            }

            SqlConnection sqlcon = new SqlConnection(con);

            sqlcon.Open();

            SqlCommand sqlcomand = new SqlCommand(sqlselect, sqlcon);
            SqlDataReader sqlread = sqlcomand.ExecuteReader();

            sqlcon.Close();

            sqlcon.Open();

            sqlcomand = new SqlCommand(sqlselect, sqlcon);
            sqlread = sqlcomand.ExecuteReader();

            while (sqlread.Read())
            {
                WeatherList weather = new WeatherList();

                weather.date = sqlread.GetString(1);
                weather.time = sqlread.GetString(2);
                weather.T = sqlread.GetString(3);
                weather.humidity = sqlread.GetString(4);
                weather.Td = sqlread.GetString(5);
                weather.pressure = sqlread.GetString(6);
                weather.direction = sqlread.GetString(7);
                weather.speed = sqlread.GetString(8);
                weather.sky = sqlread.GetString(9);
                weather.h = sqlread.GetString(10);
                weather.VV = sqlread.GetString(11);
                weather.other = sqlread.GetString(12);

                weatherLists.Add(weather);

            }

            //добавление строк в html файле
            ViewBag.People = new List<WeatherList>(weatherLists);
            sqlcon.Close();
            return View();
        }


        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}