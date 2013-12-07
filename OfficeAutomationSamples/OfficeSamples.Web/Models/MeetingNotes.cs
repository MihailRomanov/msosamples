using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Xml.Serialization;

namespace OfficeSamples.Models
{
	[XmlRoot("meetingNotes", Namespace = "urn:MeetingNotes")]
	public class MeetingNotes
	{
		public MeetingNotes()
		{
			Participants = new List<Participant>();
			Decisions = new List<Decision>();
		}

		[XmlAttribute("subject")]
		[Required]
		public string Subject { get; set; }

		[XmlAttribute("date")]
		[Required]
		public DateTime Date { get; set; }

		[XmlAttribute("secretary")]
		[Required]
		public string Secretary { get; set; }

		[XmlArray("participants")]
		public List<Participant> Participants { get; set; }

		[XmlArray("decisions")]
		public List<Decision> Decisions { get; set; }

	}
}