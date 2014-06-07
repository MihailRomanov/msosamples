using System;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace OpenXMLSdkSamples
{
	[TestClass]
	public class MetadataSamples
	{
		[TestMethod]
		public void TestMethod1()
		{
			const string fileName = @"TestDocuments\MetadataSample.docx";

			using (var document = WordprocessingDocument.Open(fileName, false))
			{
				var extendedProperties = document.ExtendedFilePropertiesPart.Properties;

				Console.WriteLine("Application : {0}", extendedProperties.Application.Text);
				Console.WriteLine("ApplicationVersion : {0}", extendedProperties.ApplicationVersion.Text);
				Console.WriteLine("Characters : {0}", extendedProperties.Characters.Text);
				Console.WriteLine("CharactersWithSpaces : {0}", extendedProperties.CharactersWithSpaces.Text);
				Console.WriteLine("Company : {0}", extendedProperties.Company.Text);
				Console.WriteLine("HyperlinksChanged : {0}", extendedProperties.HyperlinksChanged.Text);
				Console.WriteLine("Lines : {0}", extendedProperties.Lines.Text);
				Console.WriteLine("LinksUpToDate : {0}", extendedProperties.LinksUpToDate.Text);
				Console.WriteLine("Pages : {0}", extendedProperties.Pages.Text);
				Console.WriteLine("Paragraphs : {0}", extendedProperties.Paragraphs.Text);
				Console.WriteLine("ScaleCrop : {0}", extendedProperties.ScaleCrop.Text);
				Console.WriteLine("SharedDocument : {0}", extendedProperties.SharedDocument.Text);
				Console.WriteLine("Template : {0}", extendedProperties.Template.Text);
				Console.WriteLine("TotalTime : {0}", extendedProperties.TotalTime.Text);
				Console.WriteLine("Words : {0}", extendedProperties.Words.Text);
			}
		}
	}
}
