using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace GenerateWordDocument
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
		public string Subject { get; set; }
		
		[XmlAttribute("date")]
		public DateTime Date { get; set; }
		
		[XmlAttribute("secretary")]
		public string Secretary { get; set; }
		
		[XmlArray("participants")]
		public List<Participant> Participants { get; set; }
		
		[XmlArray("decisions")]
		public List<Decision> Decisions { get; set; }
	}

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

	[XmlType("participant")]
	public class Participant
	{
		[XmlAttribute("name")]
		public string Name { get; set; }
	}
}
