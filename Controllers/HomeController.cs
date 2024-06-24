using Microsoft.AspNetCore.Mvc;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Diagnostics;
using WebApplication1.Models;

namespace WebApplication1.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {

            // 创建一个新的工作簿
            IWorkbook workbook = new XSSFWorkbook();

            // 创建一个新的工作表
            ISheet sheet = workbook.CreateSheet("Sheet1");

            // 创建第一行并写入数据
            IRow row1 = sheet.CreateRow(0);
            row1.CreateCell(0).SetCellValue("ID");
            row1.CreateCell(1).SetCellValue("Name");
            row1.CreateCell(2).SetCellValue("Age");

            // 创建第二行并写入数据
            IRow row2 = sheet.CreateRow(1);
            row2.CreateCell(0).SetCellValue(1);
            row2.CreateCell(1).SetCellValue("John Doe");
            row2.CreateCell(2).SetCellValue(30);

            // 创建第三行并写入数据
            IRow row3 = sheet.CreateRow(2);
            row3.CreateCell(0).SetCellValue(2);
            row3.CreateCell(1).SetCellValue("Jane Smith");
            row3.CreateCell(2).SetCellValue(25);

            // 将工作簿写入文件
            string filePath = "example.xlsx";
            using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs);
            }
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
