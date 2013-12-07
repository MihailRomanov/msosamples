using System.Web.Mvc;

namespace OfficeSamples.Areas.GenerateWordFormSample
{
	public class GenerateWordFormSampleAreaRegistration : AreaRegistration
	{
		public override string AreaName
		{
			get
			{
				return "GenerateWordFormSample";
			}
		}

		public override void RegisterArea(AreaRegistrationContext context)
		{
			context.MapRoute(
				"GenerateWordFormSample_default",
				"GenerateWordFormSample/{controller}/{action}/{id}",
				new { action = "Index", id = UrlParameter.Optional },
				namespaces: new[] { "OfficeSamples.Areas.GenerateWordFormSample.Controllers" }
			);
		}
	}
}