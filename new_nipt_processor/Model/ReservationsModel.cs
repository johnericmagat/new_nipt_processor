using System;

namespace new_nipt_processor.Model
{
	public class ReservationsModel
	{
		public Int32 Id { get; set; }
		public String Illumina_Report_Id { get; set; }
		public DateTime Reserve_Datetime { get; set; }
		public String Name { get; set; }
		public String Email { get; set; }
		public DateTime Created_At { get; set; }
	}
}
