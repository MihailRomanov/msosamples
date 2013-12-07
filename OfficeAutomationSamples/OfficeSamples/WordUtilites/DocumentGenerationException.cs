using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace OfficeSamples.WordUtilites
{
	public class DocumentGenerationException : Exception
	{
		public DocumentGenerationException()
		{ }

		public DocumentGenerationException(string message)
			: base(message)
		{ }
	}
}