using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ExcelDataReader;
using LaborReport.Datas;
using LaborReport.Datas.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;

namespace LaborReport.Controllers
{
    public class InOutsController : Controller
    {
        private readonly ApplicationDbContext _context;

        public InOutsController(ApplicationDbContext context)
        {
            _context = context;
        }
        public IActionResult Index()
        {
            var model = new IndexViewModel();
            model.Table = new Data[7, 7];
            var month = DateTime.Now.Month;
            var year = DateTime.Now.Year;
            var days = DateTime.DaysInMonth(year, month);

            var startDate = new DateTime(year, month, 1);
            var endDate = new DateTime(year, month, days);

            var inOuts = _context.InOuts.Where(x => x.Time >= startDate && x.Time <= endDate);

            for (int day = 1; day <= days; day++)
            {
                var date = new DateTime(year, month, day);
                var col = (int)date.DayOfWeek > 0 ? (int)date.DayOfWeek - 1 : (int)date.DayOfWeek + 6;
                var row = (day - col - 2) / 7 + 1;

                var data = new Data()
                {
                    DateString = date.ToString("dd/MM"),
                    Date = date,
                    LinkData = date.ToString("yyyy-MM-dd")
                };
                if (inOuts.FirstOrDefault(x => x.Time.Date == date.Date) != null)
                    data.HasData = true;
                else
                    data.HasData = false;

                model.Table[row, col] = data;
                model.RowMax = row;
            }

            return View(model);
        }

        public IActionResult Detail(string id)
        {
            DateTime date;
            List<DetailViewModel> model;
            try
            {
                date = DateTime.Parse(id);
                model = GetData(date.Date, date.Date);
            }
            catch
            {
                return BadRequest();
            }
            return View(model);
        }

        private List<DetailViewModel> GetData(DateTime start, DateTime end)
        {
            var stacks = new List<Stack>();
            var result = new List<DetailViewModel>();
            var data = _context.InOuts.Where(o => o.Time >= start.AddDays(-5) && o.Time <= end.AddDays(5)).OrderBy(o => o.Time).ToList();
            for (int i = 0; i < data.Count; i++)
            {
                if (data[i].Event.Contains("Vào"))
                {
                    if (stacks.FirstOrDefault(x => x.CardNumber == data[i].CardNumber) == null)
                    {
                        var stack = new Stack { CardNumber = data[i].CardNumber, Time = data[i].Time };
                        stacks.Add(stack);
                        if (result.FirstOrDefault(x => x.CardNumber == data[i].CardNumber) == null)
                        {
                            result.Add(new DetailViewModel { CardNumber = data[i].CardNumber });
                        }
                    }
                }
                else
                {
                    var stack = stacks.FirstOrDefault(x => x.CardNumber == data[i].CardNumber);
                    if (stack != null)
                    {
                        if (data[i].Time.Date >= start && data[i].Time.Date <= end)
                        {
                            var item = result.FirstOrDefault(x => x.CardNumber == data[i].CardNumber);
                            if (item != null)
                            {
                                TimeSpan diff = data[i].Time.Subtract(stack.Time);
                                item.TotalTime = item.TotalTime + diff.TotalSeconds / 60;
                            }
                        }

                        stacks.Remove(stack);
                    }
                }
            }

            return result.Where(x => x.TotalTime != 0).ToList();
        }

        public IActionResult Upload()
        {
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Upload(FileUpload file)
        {
            if (_context.InOuts.FirstOrDefault(x => x.Time.Date == file.UploadDate) != null)
            {
                ModelState.AddModelError("", "Ngày đã chọn đã có dữ liệu trên hệ thống, không thể upload thêm dữ liệu");
                return View(file);
            }
            DataSet result;
            bool hasData = false;
            var fileName = Guid.NewGuid().ToString();
            var filePath = $"{Directory.GetCurrentDirectory()}\\UploadData\\{fileName}.xls";

            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                await file.FormFile.CopyToAsync(stream);
            }

            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    do
                    {
                        while (reader.Read())
                        {
                        }
                    } while (reader.NextResult());

                    result = reader.AsDataSet();
                }

            }

            foreach (DataRow line in result.Tables[0].Rows)
            {
                var item = new InOut();
                try
                {
                    item.Time = DateTime.Parse(line.ItemArray[0].ToString());
                    if (item.Time.Date == file.UploadDate)
                    {
                        item.Id = Guid.NewGuid();
                        item.CardNumber = line.ItemArray[1].ToString();
                        item.Event = line.ItemArray[2].ToString();
                        item.Status = line.ItemArray[3].ToString();
                        item.Description = line.ItemArray[4].ToString();

                        hasData = true;
                        _context.InOuts.Add(item);
                    }
                }
                catch { }
            }

            if (hasData)
            {
                _context.SaveChanges();
                return RedirectToAction("Index");
            }
            else
            {
                ModelState.AddModelError("", "File upload không chứa dữ liệu của ngày lựa chọn!");
                return View(file);
            }
        }
    }

    public class DetailViewModel
    {
        public string CardNumber { get; set; }
        public double TotalTime { get; set; } = 0;
    }

    public class Stack
    {
        public string CardNumber { get; set; }
        public DateTime Time { get; set; }
    }

    public class FileUpload
    {
        [Required]
        [Display(Name = "Ngày upload dữ liệu")]
        public DateTime UploadDate { get; set; }

        [Required]
        [Display(Name = "Chọn tệp tin dữ liệu")]
        public IFormFile FormFile { get; set; }
    }

    public class Data
    {
        public string DateString { get; set; }
        public bool HasData { get; set; }
        public string LinkData { get; set; }
        public DateTime Date { get; set; }
    }

    public class IndexViewModel
    {
        public Data[,] Table { get; set; }
        public int RowMax { get; set; } = 0;
    }
}