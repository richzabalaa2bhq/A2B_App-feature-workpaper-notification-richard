
using A2B_App.Server.Services;
using A2B_App.Shared.Sox;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;


namespace A2B_App.Server.Controllers
{



    [Route("api/[controller]")]
    [ApiController]
    public class FileDownloadController : ControllerBase
    {

        private readonly IWebHostEnvironment _environment;
        private readonly IConfiguration _config;

        public FileDownloadController(IWebHostEnvironment environment, IConfiguration config)
        {
            _environment = environment;
            _config = config;
        }


        [AllowAnonymous]
        [HttpGet("download/{filename}")]
        public async Task<IActionResult> GetDownloadAsync(string filename)
        {
            //test link = https://localhost:44344/SampleSelection/download/Draft-TestRound-20200408_183306.xlsx

            try
            {
                string startupPath = Directory.GetCurrentDirectory();
                string path = Path.Combine(startupPath, "include", "sampleselection", "download", filename);

                var memory = new MemoryStream();
                using (var stream = new FileStream(path, FileMode.Open))
                {
                    await stream.CopyToAsync(memory);
                }
                memory.Position = 0;
                var ext = Path.GetExtension(path).ToLowerInvariant();

                return File(memory, GetMimeTypes()[ext], Path.GetFileName(path));
            }
            catch (Exception ex)
            {
                FileLog.Write($"Error GetDownloadAsync {ex}", "ErrorGetDownloadAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetDownloadAsync");
                if (_environment.IsDevelopment())
                {
                    return BadRequest("Oops something when wrong!." + ex.ToString());
                }
                else
                {
                    return BadRequest("Oops something when wrong!.");
                }

            }
        }


        [AllowAnonymous]
        [HttpGet("sod/{filename}")]
        public async Task<IActionResult> GetSodFileAsync(string filename)
        {
            //test link = https://localhost:44344/SampleSelection/download/Draft-TestRound-20200408_183306.xlsx

            try
            {
                string startupPath = Directory.GetCurrentDirectory();
                string path = Path.Combine(startupPath, "include", "sod", filename);

                var memory = new MemoryStream();
                using (var stream = new FileStream(path, FileMode.Open))
                {
                    await stream.CopyToAsync(memory);
                }
                memory.Position = 0;
                var ext = Path.GetExtension(path).ToLowerInvariant();

                return File(memory, GetMimeTypes()[ext], Path.GetFileName(path));
            }
            catch (Exception ex)
            {
                FileLog.Write($"Error GetSodFileAsync {ex}", "ErrorGetSodFileAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetSodFileAsync");
                if (_environment.IsDevelopment())
                {
                    return BadRequest("Oops something when wrong!." + ex.ToString());
                }
                else
                {
                    return BadRequest("Oops something when wrong!.");
                }

            }
        }

        [AllowAnonymous]
        [HttpGet("sod/template/{filename}")]
        public async Task<IActionResult> GetSodFileTemplateAsync(string filename)
        {

            try
            {
                string startupPath = Directory.GetCurrentDirectory();
                string path = Path.Combine(startupPath, "include", "sod", "template", filename);

                var memory = new MemoryStream();
                using (var stream = new FileStream(path, FileMode.Open))
                {
                    await stream.CopyToAsync(memory);
                }
                memory.Position = 0;
                var ext = Path.GetExtension(path).ToLowerInvariant();

                return File(memory, GetMimeTypes()[ext], Path.GetFileName(path));
            }
            catch (Exception ex)
            {
                FileLog.Write($"Error GetSodFileAsync {ex}", "ErrorGetSodFileAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetSodFileAsync");
                if (_environment.IsDevelopment())
                {
                    return BadRequest("Oops something when wrong!." + ex.ToString());
                }
                else
                {
                    return BadRequest("Oops something when wrong!.");
                }

            }
        }


        private Dictionary<string, string> GetMimeTypes()
        {
            return new Dictionary<string, string>
            {
                {".txt", "text/plain"},
                {".pdf", "application/pdf"},
                {".doc", "application/vnd.ms-word"},
                {".docx", "application/vnd.ms-word"},
                {".xls", "application/vnd.ms-excel"},
                {".xlsx", "application/vnd.ms-excel"},
                {".xlsm", "application/vnd.ms-excel.addin.macroEnabled.12"},
                {".png", "image/png"},
                {".jpg", "image/jpeg"},
                {".mp4", "video/mp4"},
            };
        }


    }

    

}
