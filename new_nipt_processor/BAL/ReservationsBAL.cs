using new_nipt_processor.DAL;
using System.Data;

namespace new_nipt_processor.BAL
{
	public class ReservationsBAL
	{
		public static DataTable FilterReservations()
		{
			return ReservationsDAL.FilterReservations("FilterReservations");
		}

		public static DataTable FilterReservationsAll()
		{
			return ReservationsDAL.FilterReservationsAll("FilterReservationsAll");
		}
	}
}
