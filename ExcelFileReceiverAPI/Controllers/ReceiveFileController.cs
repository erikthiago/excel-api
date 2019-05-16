using ExcelFileReceiverAPI.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.IO;
using System.Threading.Tasks;

namespace ExcelFileReceiverAPI.Controllers
{
    [Route("api/receivefile")]
    [ApiController]
    public class ReceiveFileController : ControllerBase
    {
        Teste teste = new Teste();

        [HttpPost]
        public async Task<IActionResult> UploadFile(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                return RedirectToAction("");
            }

            using (var memoryStream = new MemoryStream())
            {
                await file.CopyToAsync(memoryStream).ConfigureAwait(false);

                using (var package = new ExcelPackage(memoryStream))
                {
                    for (int i = 1; i <= package.Workbook.Worksheets.Count; i++)
                    {
                        var totalRows = package.Workbook.Worksheets[i].Dimension?.Rows;
                        var totalCollumns = package.Workbook.Worksheets[i].Dimension?.Columns;
                        for (int j = 1; j <= totalRows.Value; j++)
                        {
                            for (int k = 1; k <= totalCollumns.Value; k++)
                            {
                                teste.T1 = package.Workbook.Worksheets[i].Cells[j, k].Value.ToString();
                                teste.T2 = package.Workbook.Worksheets[i].Cells[j, k].Value.ToString();
                            }
                        }
                    }

                    return Content(teste.T1);
                }
            }
        }
    }
}