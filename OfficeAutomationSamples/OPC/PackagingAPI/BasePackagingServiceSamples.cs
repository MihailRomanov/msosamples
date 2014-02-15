using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace PackagingAPI
{
	[TestClass]
	public class BasePackagingServiceSamples
	{
		[TestMethod]
		public void CreateWordDocumentWithTextAndImage()
		{
			// Part Names
			const string mainDocumentPartName = "/document/main.xml";
			const string imagePartName = "/images/cat.jpeg";

			// Content types
			const string mainDocumentPartContentType =
				"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
			const string imagePartContentType =
				System.Net.Mime.MediaTypeNames.Image.Jpeg;

			// Relationsip Types
			const string mainDocumentPartRelationshipType =
				"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
			const string imagePartRelationshipType = 
				"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

			// Image relationship Id
			const string imageRelationshipId = "rId1";

			// Create new Word document
			using (var document = Package.Open("result.docx", System.IO.FileMode.Create))
			{
				// Add main part
				var mainPartUri = new Uri(mainDocumentPartName, UriKind.Relative);
				var mainPart = document.CreatePart(mainPartUri, mainDocumentPartContentType);

				using (var mainPartStream = mainPart.GetStream())
				{
					mainPartStream.Write(Properties.Resources.main, 0, Properties.Resources.main.Length);
					mainPartStream.Close();
				}

				document.CreateRelationship(mainPartUri, TargetMode.Internal, mainDocumentPartRelationshipType);

				// Add image
				var imagePartUri = new Uri(imagePartName, UriKind.Relative);
				var imagePart = document.CreatePart(imagePartUri, imagePartContentType);

				using (var imagePartStream = imagePart.GetStream())
				{
					imagePartStream.Write(Properties.Resources.cat, 0, Properties.Resources.cat.Length);
					imagePartStream.Close();
				}

				// Create relative uri for image
				var relativeUri = PackUriHelper.GetRelativeUri(mainPartUri, imagePartUri);
				mainPart.CreateRelationship(relativeUri, TargetMode.Internal, imagePartRelationshipType, imageRelationshipId);

				document.Close();
			}
		}

		[TestMethod]
		public void ReadWordDocumentImages()
		{
			// Relationsip Types
			const string mainDocumentPartRelationshipType =
				"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
			const string imagePartRelationshipType =
				"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

			CreateWordDocumentWithTextAndImage();

			using (var document = Package.Open("result.docx", System.IO.FileMode.Open))
			{
				var mainPartRelationship = document.GetRelationshipsByType(mainDocumentPartRelationshipType).Single();
				
				var mainPartName = PackUriHelper.ResolvePartUri(new Uri("/", UriKind.Relative), mainPartRelationship.TargetUri);
				var mainPart = document.GetPart(mainPartName);

				foreach (var imageRelationship in mainPart.GetRelationshipsByType(imagePartRelationshipType))
				{
					var imagePartName = PackUriHelper.ResolvePartUri(mainPartName, imageRelationship.TargetUri);
					var imagePart = document.GetPart(imagePartName);

					var fileName = Path.GetFileName(imagePartName.OriginalString);

					using (var file = new FileStream(fileName, FileMode.Create))
					{
						var imageStream = imagePart.GetStream();
						imageStream.CopyTo(file);

						file.Close();
						imageStream.Close();
					}
				}
			}
		}

		[TestMethod]
		public void ChangeWordDocumentMainPart()
		{
			const string mainDocumentPartRelationshipType =
				"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

			CreateWordDocumentWithTextAndImage();

			using (var document = Package.Open("result.docx", System.IO.FileMode.Open))
			{
				var mainPartRelationship = document.GetRelationshipsByType(mainDocumentPartRelationshipType).Single();

				var mainPartName = PackUriHelper.ResolvePartUri(new Uri("/", UriKind.Relative), mainPartRelationship.TargetUri);
				var mainPart = document.GetPart(mainPartName);

				using (var mainPartStream = mainPart.GetStream())
				{
					mainPartStream.SetLength(0);
					mainPartStream.Write(Properties.Resources.main2, 0, Properties.Resources.main2.Length);

					mainPartStream.Close();
				}

				document.Close();
			}

		}
	}
}
