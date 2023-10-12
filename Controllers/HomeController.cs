using DocumentSecurity.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace DocumentSecurity.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult EncryptDocument()
        {
            FileStream fileStream = new FileStream(Path.GetFullPath("Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            WordDocument document = new WordDocument(fileStream, FormatType.Automatic);
            document.EncryptDocument("syncfusion");
            MemoryStream ms = new MemoryStream();
            document.Save(ms, FormatType.Docx);
            document.Close();
            return File(ms, "application/msword", "Sample.docx");
        }

        public IActionResult DecryptDocument()
        {
            FileStream fileStream = new FileStream(Path.GetFullPath("Data/Sample.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            WordDocument document = new WordDocument(fileStream, FormatType.Docx, "syncfusion");
            IWParagraph paragraph = document.LastSection.AddParagraph();
            IWTextRange text = paragraph.AppendText("\nDemo For Document Decryption with Essential DocIO");
            text.CharacterFormat.FontSize = 16f;
            text.CharacterFormat.FontName = "Bitstream Vera Serif";
            text = paragraph.AppendText("\nThis document is Decrypted");
            text.CharacterFormat.FontSize = 16f;
            text.CharacterFormat.FontName = "Bitstream Vera Serif";
            MemoryStream ms = new MemoryStream();
            document.Save(ms, FormatType.Docx);
            document.Close();
            return File(ms, "application/msword", "Decrypt.docx");
        }
        public IActionResult Index()
        {
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