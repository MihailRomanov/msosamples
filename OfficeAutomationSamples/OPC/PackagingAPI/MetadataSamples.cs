using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace PackagingAPI
{
	[TestClass]
	public class MetadataSamples
	{
		[TestMethod]
		public void DirectReadProperties()
		{
			const string corePropertiesUri = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";

			using (var package = Package.Open(@"TestDocuments\TestDocument.docx", FileMode.Open))
			{
				var corePropertiesRelationship = package.GetRelationshipsByType(corePropertiesUri)
					.Single();
				var corePropertiesPart = package.GetPart(PackUriHelper.CreatePartUri(corePropertiesRelationship.TargetUri));

				var reader = new StreamReader(corePropertiesPart.GetStream());
				Console.WriteLine(reader.ReadToEnd());
			}
		}

		[TestMethod]
		public void ShowPackageProperties()
		{
			using (var package = Package.Open(@"TestDocuments\TestDocument.docx", FileMode.Open))
			{
				var properties = package.PackageProperties;

				Console.WriteLine("Category : {0}", properties.Category);
				Console.WriteLine("ContentStatus : {0}", properties.ContentStatus);
				Console.WriteLine("ContentType : {0}", properties.ContentType);
				Console.WriteLine("Created : {0}", properties.Created);
				Console.WriteLine("Creator : {0}", properties.Creator);
				Console.WriteLine("Description : {0}", properties.Description);
				Console.WriteLine("Identifier : {0}", properties.Identifier);
				Console.WriteLine("Keywords : {0}", properties.Keywords);
				Console.WriteLine("Language : {0}", properties.Language);
				Console.WriteLine("LastModifiedBy : {0}", properties.LastModifiedBy);
				Console.WriteLine("LastPrinted : {0}", properties.LastPrinted);
				Console.WriteLine("Modified : {0}", properties.Modified);
				Console.WriteLine("Revision : {0}", properties.Revision);
				Console.WriteLine("Subject : {0}", properties.Subject);
				Console.WriteLine("Title : {0}", properties.Title);
				Console.WriteLine("Version : {0}", properties.Version);
			}
		}

		[TestMethod]
		public void ChangePackageProperties()
		{
			File.Copy(@"TestDocuments\TestDocument.docx", @"TestDocuments\TestDocument2.docx", true);

			using (var package = Package.Open(@"TestDocuments\TestDocument2.docx", FileMode.Open, FileAccess.ReadWrite))
			{
				var properties = package.PackageProperties;

				properties.Title = "Changed document";
				properties.Modified = DateTime.Now;
				properties.Description = string.Format("Document changed at {0}", properties.Modified);

				package.Close();
			}

			System.Diagnostics.Process.Start(@"TestDocuments\TestDocument2.docx");
		}
	}
}
