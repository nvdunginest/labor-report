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
            var stacks = new List<Stack>();
            var result = new List<IndexViewModel>();
            var data = _context.InOuts.OrderBy(o => o.Time).ToList();
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
                            result.Add(new IndexViewModel { CardNumber = data[i].CardNumber });
                        }
                    }
                }
                else
                {
                    var stack = stacks.FirstOrDefault(x => x.CardNumber == data[i].CardNumber);
                    if (stack != null)
                    {
                        var item = result.FirstOrDefault(x => x.CardNumber == data[i].CardNumber);
                        if (item != null)
                        {
                            TimeSpan diff = data[i].Time.Subtract(stack.Time);
                            item.TotalTime = item.TotalTime + diff.TotalSeconds / 60;
                        }

                        stacks.Remove(stack);
                    }
                }
            }

            return View(result.OrderBy(x => x.CardNumber));
        }

        public IActionResult Upload()
        {
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Upload(FileUpload file)
        {
            DataSet result;
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
                    item.Id = Guid.NewGuid();
                    item.CardNumber = line.ItemArray[1].ToString();
                    item.Event = line.ItemArray[2].ToString();
                    item.Status = line.ItemArray[3].ToString();
                    item.Description = line.ItemArray[4].ToString();

                    _context.InOuts.Add(item);
                }
                catch { }
            }

            _context.SaveChanges();
            return RedirectToAction("Index");
        }
    }

    public class IndexViewModel
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
        [Display(Name = "File")]
        public IFormFile FormFile { get; set; }
    }
}