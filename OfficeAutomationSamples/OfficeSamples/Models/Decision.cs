using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Xml.Serialization;

namespace OfficeSamples.Models
{
	[XmlType("decision")]
	public class Decision
	{
		[XmlAttribute("problem")]
		public string Problem { get; set; }

		[XmlAttribute("solution")]
		public string Solution { get; set; }

		[XmlAttribute("responsible")]
		public string Responsible { get; set; }

		[XmlAttribute("controlDate")]
		public DateTime ControlDate { get; set; }
	}
}