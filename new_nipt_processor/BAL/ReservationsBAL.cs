using new_nipt_processor.DAL;
using System.Data;

namespace new_nipt_processor.BAL
{
	public class ReservationsBAL
	{
		public static DataTable FilterUsers(string dateStart, string dateEnd)
		{
			return ReservationsDAL.FilterReservations("FilterReservations", dateStart, dateEnd);
		}

		public static DataTable FilterReservationsAll()
		{
			return ReservationsDAL.FilterReservationsAll("FilterReservationsAll");
		}
	}
}
