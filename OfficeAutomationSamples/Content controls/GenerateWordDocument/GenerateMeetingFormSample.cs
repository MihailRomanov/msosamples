using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Serialization;
using DocumentFormat.OpenXml.CustomXmlDataProperties;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace GenerateWordDocument
{
	[TestClass]
	public class GenerateMeetingFormSample
	{
		[TestMethod]
		public void GenerateDocument()
		{
			const string templateName = @"Template.docx";
			const string resultDocumentName = @"MeetingNotes.docx";

			var meetingNotes = new MeetingNotes()
			{
				Subject = "Result of meeting note tamplates development",
				Date = new DateTime(2013, 09, 21),
				Secretary = "Romanov M.",
				Participants = new List<Participant>
				{
					new Participant { Name = "Romanov M." }, 
					new Participant { Name = "Bushmelev S." }, 
					new Participant { Name = "Prosalov M." } 
				},
				Decisions = new List<Decision>
				{
					new Decision { Problem = "What result?" , Solution = "Accepted!", 
						ControlDate = new DateTime(2013, 09, 23), Responsible = "Romanov M."},

					new Decision { Problem = "What we will do next time?" , Solution = "Introduce!!!", 
						ControlDate = new DateTime(2013, 12, 31), Responsible = "Romanov M."}
				}
			};

			var serializer = new XmlSerializer(typeof(MeetingNotes));
			var serializedDataStream = new MemoryStream();

			var namespaces = new XmlSerializerNamespaces();
			namespaces.Add("", "");

			serializer.Serialize(serializedDataStream, meetingNotes, namespaces);
			serializedDataStream.Seek(0, SeekOrigin.Begin);

			File.Copy(templateName, resultDocumentName, true);

			using (var document = WordprocessingDocument.Open(resultDocumentName, true))
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
