using OfficeSamples.Models;
using OfficeSamples.WordUtilites;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace OfficeSamples.Areas.GenerateWordFormSample.Controllers
{
	public class HomeController : Controller
	{
		public HomeController()
		{
			ViewBag.Title = "Word Form generation sample";
		}

		public ActionResult Index()
		{
			return View();
		}

		[HttpPost]
		public JsonResult GenerateReport(MeetingNotes meetingNotes)
		{
			GenerationResult result = new GenerationResult();

			if (ModelState.IsValid)
			{
				try
				{
					var generator = new WordFormReportGenerator();
					var template = new MemoryStream(Properties.Resources.Template1);

					var document = generator.GenerateDocument(template, meetingNotes) as MemoryStream;

					var id = Guid.NewGuid();
					Session[id.ToString()] = document.ToArray();
					result.DocumentId = id;
				}
				catch (DocumentGenerationException ex)
				{
					result.Error = string.Join("", ex.Message);
				}
			}
			else
			{
				result.Error = string.Join("", ModelState.Values.SelectMany(v => v.Errors.Select(err => err.ErrorMessage)));
			}

			return new JsonResult() { Data = result };
		}

		public ActionResult GetDocument(Guid documentId)
		{
			var document = Session[documentId.ToString()] as byte[];
			
			if (document != null)
			{
				return File(document, "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
			}
			else
				return new HttpNotFoundResult();
		}
	}
}