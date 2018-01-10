using System.IO;
using System.Speech.Synthesis;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using WebApplication1.Models;

namespace WebApplication1.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public async Task<ActionResult> Index(HttpPostedFileBase file)
        {
            if (Validator.IsValidFormat(file))
            {
                string path = HttpContext.Server.MapPath("~/App_Data/");
                await TypeConverter.FileToWav(file, path);            
            }
            else
            {
                ModelState.AddModelError("", "You must upload only file with .pdf or .txt extension");
                return View();
            }
            return RedirectToAction("Result", "Home", new { arg = Path.GetFileNameWithoutExtension(file.FileName) });
        }

        public ActionResult Result(string arg)
        {
            return View((object)arg);
        }

        public FileResult Download(string fileName)
        {
            if (fileName == null)
            {
                RedirectToAction("Index", "Home");
            }
            string path = HttpContext.Server.MapPath("~/App_Data/" + fileName);
            byte[] fileBytes = System.IO.File.ReadAllBytes(path);
            return File(fileBytes, "wav", fileName);
        }
    }
}