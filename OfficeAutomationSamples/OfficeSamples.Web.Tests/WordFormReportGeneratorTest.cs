using System;
using OfficeSamples.WordUtilites;
using DocumentFormat.OpenXml.Packaging;
using System.IO;
using DocumentFormat.OpenXml;
using System.Linq;
using System.Xml.Linq;
using OfficeSamples.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace OfficeSamples.Tests
{
	[TestClass]
	public class WordFormReportGeneratorTest
	{
		[TestMethod]
		public void GetAppropriateXmlPartSuccessTest()
		{
			var generator = new WordFormReportGenerator();
			var template = WordprocessingDocument.Open(new MemoryStream(Properties.Resources.MeetingNotesTemplate), false);
			var serializedData = Properties.Resources.CorrectMeetingNotesData;

			var part = generator.GetAppropriateXmlPart(template, new MemoryStream(serializedData, false));
			Assert.IsNotNull(part);

			Assert.AreEqual(serializedData.Length, part.GetStream().Length);
		}

		[TestMethod]
		[ExpectedException(typeof(DocumentGenerationException))]
		public void GetAppropriateXmlPartFaildTest()
		{
			var generator = new WordFormReportGenerator();
			var template = WordprocessingDocument.Open(new MemoryStream(Properties.Resources.MeetingNotesTemplate), false);
			var serializedData = Properties.Resources.NoCorrectMeetingNotesData;

			var part = generator.GetAppropriateXmlPart(template, new MemoryStream(serializedData, false));

			Assert.Fail();
		}

		[TestMethod]
		public void GenerateDocumentFromTemplateStreamAndSerializedDataTest()
		{
			var generator = new WordFormReportGenerator();
			var template = new MemoryStream(Properties.Resources.MeetingNotesTemplate, false);
			var serializedData = Properties.Resources.CorrectMeetingNotesData2;

			var documentStream = generator.GenerateDocument(template, serializedData);

			var document = WordprocessingDocument.Open(documentStream, false);
			var xmlPart = document.MainDocumentPart.CustomXmlParts.First();

			Assert.AreEqual(serializedData.Length, xmlPart.GetStream().Length);
		}

		private	void CheckGenerationResult(Stream documentStream, string subject)
		{
			var document = WordprocessingDocument.Open(documentStream, false);
			var xml = XDocument.Load(document.MainDocumentPart.CustomXmlParts.First().GetStream());
			var documentSubject = xml.Descendants()
				.Single(elem => elem.Name.LocalName == "meetingNotes")
				.Attribute("subject").Value;

			Assert.AreEqual(subject, documentSubject);
		}

		[TestMethod]
		public void GenerateDocumentFromFileAndObjectTest()
		{
			var fileName = Path.GetTempFileName();

			try
			{
				var file = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite);

				var templateBody = Properties.Resources.MeetingNotesTemplate;
				file.Write(templateBody, 0, templateBody.Length);
				file.Close();

				var subject = "Sample subject";
				var meetingNotes = new MeetingNotes { Subject = subject };

				var generator = new WordFormReportGenerator();

				var documentStream = generator.GenerateDocument(fileName, meetingNotes);
				CheckGenerationResult(documentStream, subject);
			}
			finally
			{
				File.Delete(fileName);
			}
		}

		[TestMethod]
		public void GenerateDocumentFromStreamAndObjectTest()
		{
			var subject = "Sample subject";
			var meetingNotes = new MeetingNotes { Subject = subject };

			var generator = new WordFormReportGenerator();
			var documentStream = generator.GenerateDocument(new MemoryStream(Properties.Resources.MeetingNotesTemplate), meetingNotes);

			CheckGenerationResult(documentStream, subject);
		}
	}
}
