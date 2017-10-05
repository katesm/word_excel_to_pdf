using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

namespace word_to_pdf.Controllers
{

    public class ValuesController : Controller
    {
        [HttpGet("~/test")]
        public IActionResult test()
        {

            var myFile = new File
            {
                FileName = "Test.docx",
                FileBytes = System.IO.File.ReadAllBytes("Test.docx")
            };

            var result = WordToPdf(myFile);

            return result;
        }

        [Route("~/wordtopdf")]
        [HttpGet]
        public IActionResult WordToPdf([FromBody] File file)
        {

            if (ModelState.IsValid && file.FileName.IndexOf(".docx") != -1)
            {
                // Save the file first

                System.IO.File.WriteAllBytes(file.FileName, file.FileBytes);
                var newName = file.FileName.Replace(".docx", ".pdf");
                var processInfo = new ProcessStartInfo
                {
                    UseShellExecute = false, // change value to false
                    FileName = "OfficeToPDF.exe",
                    Arguments = $"{file.FileName} {newName}"

                };

                Console.WriteLine("Starting child process...");
                using (var process = Process.Start(processInfo))
                {
                    process.WaitForExit(10000); //if longer than 10 seconds exit

                    if (System.IO.File.Exists(newName))
                    {
                        var filePdf = System.IO.File.ReadAllBytes(newName);
                        return File(filePdf, "application/pdf");
                    }

                }

            }



            //Something happen
            return NotFound();
        }


        [Route("~/exceltopdf")]
        [HttpGet]
        public IActionResult ExcelToPdf([FromBody] File file)
        {

            if (ModelState.IsValid && file.FileName.IndexOf(".xlsx") != -1)
            {
                // Save the file first

                System.IO.File.WriteAllBytes(file.FileName, file.FileBytes);
                var newName = file.FileName.Replace(".xlsx", ".pdf");
                var processInfo = new ProcessStartInfo
                {
                    UseShellExecute = false, // change value to false
                    FileName = "OfficeToPDF.exe",
                    Arguments = $"{file.FileName} {newName}"

                };

                Console.WriteLine("Starting child process...");
                using (var process = Process.Start(processInfo))
                {
                    process.WaitForExit(10000); //if longer than 10 seconds exit

                    if (System.IO.File.Exists(newName))
                    {
                        var filePdf = System.IO.File.ReadAllBytes(newName);
                        return File(filePdf, "application/pdf");
                    }

                }

            }



            //Something happen
            return NotFound();
        }
    }






    public class File
    {
        [Required]
        public string FileName { get; set; }
        [Required]
        public byte[] FileBytes { get; set; }
    }
}
