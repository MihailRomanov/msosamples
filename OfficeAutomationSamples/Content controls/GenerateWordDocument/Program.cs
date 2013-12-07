using DocumentFormat.OpenXml.CustomXmlDataProperties;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace GenerateWordDocument
{
	class Program
	{
		const string TemplateName = @"Template.docx";
		const string ResultDocumentName = @"MeetingNotes.docx";

		static void Main(string[] args)
		{
			var meetingNotes = new MeetingNotes()
			{
				Subject = "Результаты разработки шаблонов отчетов совещаний",
				Date = new DateTime(2013, 09, 21),
				Secretary = "Романов М.",
				Participants = new List<Participant>
				{
					new Participant { Name = "Романов М." }, 
					new Participant { Name = "Бушмелев С." }, 
					new Participant { Name = "Просалов М." } 
				},
				Decisions = new List<Decision>
				{
					new Decision { Problem = "Как результат?" , Solution = "Устраивает!", 
						ControlDate = new DateTime(2013, 09, 23), Responsible = "Романов М."},

					new Decision { Problem = "Что дальше?" , Solution = "Внедряем!!!", 
						ControlDate = new DateTime(2013, 12, 31), Responsible = "Романов М."}
				}
			};

			var serializer = new XmlSerializer(typeof(MeetingNotes));
			var serializedDataStream = new MemoryStream();

			var namespaces = new XmlSerializerNamespaces();
			namespaces.Add("", "");
			
			serializer.Serialize(serializedDataStream, meetingNotes, namespaces);
			serializedDataStream.Seek(0, SeekOrigin.Begin);
			
			File.Copy(TemplateName, ResultDocumentName, true);

			using (var document = WordprocessingDocument.Open(ResultDocumentName, true))
			{
				var xmlpart = document.MainDocumentPart.CustomXmlParts
					.Single(xmlPart =>
						xmlPart.CustomXmlPropertiesPart.DataStoreItem.SchemaReferences.OfType<SchemaReference>()
						.Any(sr => sr.Uri.Value == "urn:MeetingNotes"));

				xmlpart.FeedData(serializedDataStream);
			}
			
		}
	}
}
