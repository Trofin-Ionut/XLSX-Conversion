using System;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.XlsIO;
using System.IO;
using System.Drawing;
using System.Web;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Linq;
using XLSL_Conversion.Models;
using OfficeOpenXml.Packaging;
using Syncfusion.XlsIO.Calculate;

namespace XLSL_Conversion.Controllers
{
    struct EraserExtra
    {
        public string Storage { get; set; }
        public string MovementType { get; set; }
    }
    public class HomeController : Controller
    {
        static List<List<string>> items = new();
        static EraserExtra ee = new();
        static Dictionary<string, Data> sums = new();
        static HashSet<string> storages = new();
        static string st;
        static string? name;
        [Route("/")]
        public IActionResult Home()
        {
            if (System.IO.File.Exists(name + ".xlsx"))
            {
                System.IO.File.Delete(name + ".xlsx");
            }
            return View("ReadExcel");
        }
        [Route("ReadExcel")]
        public async Task<IActionResult> ReadExcel(IFormFile ExcelFile)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (MemoryStream stream = new MemoryStream())
            {
                await ExcelFile.CopyToAsync(stream);

                using (var package = new ExcelPackage(stream))
                {
                    name = ExcelFile.FileName;
                    name = name.Remove(name.IndexOf('.'));
                    if (package.Workbook.Worksheets.Count > 0)
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                        st = worksheet.Name;
                        var rows = worksheet.Dimension.Rows;
                        var columns = worksheet.Dimension.Columns;
                        for (int i = 2; i < rows; i++)
                        {
                            List<string> temp = new();
                            for (int j = 1; j < columns; j++)
                            {
                                if (j != 15)
                                {
                                    temp.Add(worksheet.Cells[i, j].Text.Trim());
                                    continue;
                                }
                                string d = worksheet.Cells[i, j].Text;
                                DateTime date = DateTime.ParseExact(worksheet.Cells[i, j].Text, "h\\:mm\\:ss tt", null);
                                string result = Convert.ToString(date.TimeOfDay);
                                temp.Add(Convert.ToString(result));

                            }
                            items.Add(temp);
                        }
                    }
                    System.IO.File.Delete(ExcelFile.Name);
                }
            }
            return RedirectToAction("ReplaceStorage", "Home");
        }
        public bool UnChanged(List<string> item)
        {
            if (item[4].Equals("551"))
            {
                return false;
            }
            return item[3].Equals(st);
        }
        private bool Eraser(List<string> item)  //Predicate<List<string>> example. Usefull for List<List<string>> apparently
        {
            return item[4].Equals(ee.MovementType) && item[3].Equals(ee.Storage);
        }
        private void Order(ref List<List<string>> shift)
        {
            shift = shift.OrderBy(list => Convert.ToDateTime(list[5]).Date)
                .ToList();
        }
        [Route("shift")]
        public IActionResult Shifts()
        {
            var shift1 = items.Where(list => Convert.ToDateTime(list[14]).TimeOfDay > TimeSpan.ParseExact("05:59:59", "hh\\:mm\\:ss", null)
            && Convert.ToDateTime(list[14]).TimeOfDay < TimeSpan.ParseExact("14:30:00", "hh\\:mm\\:ss", null)
            ).ToList();
            Order(ref shift1);
            var shift2 = items.Where(list => Convert.ToDateTime(list[14]).TimeOfDay > TimeSpan.ParseExact("14:29:59", "hh\\:mm\\:ss", null)
            && Convert.ToDateTime(list[14]).TimeOfDay < TimeSpan.ParseExact("22:50:00", "hh\\:mm\\:ss", null)
            ).ToList();
            Order(ref shift2);
            var shift3 = items.Where(list => Convert.ToDateTime(list[14]).TimeOfDay > TimeSpan.ParseExact("22:49:59", "hh\\:mm\\:ss", null)
            || Convert.ToDateTime(list[14]).TimeOfDay < TimeSpan.ParseExact("06:00:00", "hh\\:mm\\:ss", null)
            ).ToList();
            Order(ref shift3);

            for (int i = 0; i < shift1.Count - 1; i++)
            {
                try
                {
                    sums[shift1[i][5]].sum1[shift1[i][3]] += Convert.ToInt32(shift1[i][8].Replace(",", ""));
                }
                catch (Exception)
                {
                    try
                    {
                        sums.Add(shift1[i][5], new Data());
                        sums[shift1[i][5]].sum1.Add(shift1[i][3], Convert.ToInt32(shift1[i][8].Replace(",", "")));
                    }
                    catch (Exception)
                    {
                        sums[shift1[i][5]].sum1.Add(shift1[i][3], Convert.ToInt32(shift1[i][8].Replace(",", "")));
                    }
                }
            }

            for (int i = 0; i < shift2.Count - 1; i++)
            {
                try
                {
                    sums[shift2[i][5]].sum2[shift2[i][3]] += Convert.ToInt32(shift2[i][8].Replace(",", ""));
                }
                catch (Exception)
                {
                    try
                    {
                        sums.Add(shift1[i][5], new Data());
                        sums[shift2[i][5]].sum2.Add(shift2[i][3], Convert.ToInt32(shift2[i][8].Replace(",", "")));
                    }
                    catch (Exception)
                    {
                        sums[shift2[i][5]].sum2.Add(shift2[i][3], Convert.ToInt32(shift2[i][8].Replace(",", "")));
                    }
                }
            }

            for (int i = 0; i < shift3.Count - 1; i++)
            {
                try
                {
                    sums[shift3[i][5]].sum3[shift3[i][3]] += Convert.ToInt32(shift3[i][8].Replace(",", ""));
                }
                catch (Exception)
                {
                    try
                    {
                        sums.Add(shift3[i][5], new Data());
                        sums[shift3[i][5]].sum3.Add(shift3[i][3], Convert.ToInt32(shift3[i][8].Replace(",", "")));
                    }
                    catch (Exception)
                    {
                        sums[shift3[i][5]].sum3.Add(shift3[i][3], Convert.ToInt32(shift3[i][8].Replace(",", "")));
                    }
                }
            }
            return RedirectToAction("CreateXLSX", "Home");
        }
        [Route("storages")]
        public IActionResult ReplaceStorage()
        {
            for (int i = 0; i < items.Count; i++)
            {
                ee = new();
                bool entered = false;
                if (items[i][4].Equals("261") || items[i][4].Equals("262"))
                {
                    ee.Storage = items[i][3];
                    ee.MovementType = items[i][4];
                    entered = true;
                    storages.Add(ee.Storage);
                    try
                    {
                        sums[items[i][5]].AddStorage(items[i][3]);
                    }
                    catch (Exception) { }
                }

                if (items[i][4].Equals("551"))
                {
                    try
                    {
                        sums[items[i][5]].AddStorage(items[i][3]);
                    }
                    catch (Exception)
                    {
                        sums.TryAdd(items[i][5], new Data());
                        sums[items[i][5]].AddStorage(items[i][3]);
                    }
                    storages.Add(items[i][3]);
                }
                if (!entered)
                {
                    sums.TryAdd(items[i][5], new Data());
                    continue;
                }
                for (int k = 0; k < items.Count; k++)
                {
                    if (items[k][6].Equals(items[i][6]))
                    {
                        items[k][3] = ee.Storage;
                    }
                }
                if (ee.Storage != null && ee.MovementType != null)
                {
                    items.RemoveAll(Eraser);
                }
                entered = false;
            }
            HashSet<string> unChanged = new();
            bool exists = false;
            for (int i = 0; i < items.Count; i++)
            {
                foreach (Data dic in sums.Values)
                {
                    if (dic.sum1.ContainsKey(items[i][3]) || dic.sum2.ContainsKey(items[i][3]) || dic.sum3.ContainsKey(items[i][3]))
                    {
                        exists = true;
                        break;
                    }
                }
                if (!exists)
                {
                    unChanged.Add(items[i][3]);
                    exists = false;
                    break;
                }
                exists = false;
            }
            foreach (var storage in unChanged)
            {
                st = storage;
                items.RemoveAll(UnChanged);
            }
            unChanged = new();
            return RedirectToAction("Shifts", "Home");
        }

        [Route("xlsx")]
        public async Task<IActionResult> CreateXLSX()
        {

            ExcelPackage excel = new();
            var worksheet = excel.Workbook.Worksheets.Add(name + ".xlsx");
            worksheet.TabColor = System.Drawing.Color.Black;
            worksheet.DefaultRowHeight = 30;
            worksheet.Row(1).Height = 40;
            worksheet.Row(1).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            worksheet.Row(1).Style.Font.Bold = true;
            worksheet.Cells[2, 1].Value = "Locatie";
            int z = 3;
            foreach (var st in storages)
            {
                worksheet.Cells[z++, 1].Value = st;
            }
            int first = 2;
            int end = 4;
            sums.Keys.Order();
            foreach (var key in sums.Keys)
            {
                if (key == null)
                {
                    continue;
                }
                worksheet.Cells[1, first, 1, end].Merge = true;
                worksheet.Cells[1, first].Value = key;
                int k = 3;
                foreach (var st in storages)
                {
                    if (st == null)
                    {
                        continue;
                    }
                    worksheet.Cells[2, first].Value = "Schimbul 1";
                    worksheet.Cells[2, first + 1].Value = "Schimbul 2";
                    worksheet.Cells[2, first + 2].Value = "Schimbul 3";
                    if (sums[key].sum1.ContainsKey(st))
                    {
                        worksheet.Cells[k, first].Value = sums[key].sum1[st];
                    }
                    else
                    {
                        worksheet.Workbook.Worksheets[name + ".xlsx"].SetValue(k, first, 1);
                    }
                    if (sums[key].sum2.ContainsKey(st))
                    {
                        worksheet.Cells[k, first + 1].Value = sums[key].sum2[st];
                    }
                    else
                    {
                        worksheet.Workbook.Worksheets[name + ".xlsx"].SetValue(k, first + 1, 1);
                    }
                    if (sums[key].sum3.ContainsKey(st))
                    {
                        worksheet.Cells[k, first + 2].Value = sums[key].sum3[st];
                    }
                    else
                    {
                        worksheet.Workbook.Worksheets[name + ".xlsx"].SetValue(k, first + 2, 1);
                    }
                    k++;
                }
                first += 3;
                end += 3;
            }
            worksheet.Columns.AutoFit();
            string fileName = name + ".xlsx";
            var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            await excel.SaveAsync();
            items.Clear();
            storages.Clear();
            st = "";
            sums.Clear();
            return File(excel.GetAsByteArray(), contentType, fileName);
        }
    }
}

