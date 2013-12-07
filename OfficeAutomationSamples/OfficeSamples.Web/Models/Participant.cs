using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;
using System.Xml.Serialization;

namespace OfficeSamples.Models
{
	[XmlType("participant")]
	public class Participant
	{
		[XmlAttribute("name")]
		[Required]
		public string Name { get; set; }
	}
}