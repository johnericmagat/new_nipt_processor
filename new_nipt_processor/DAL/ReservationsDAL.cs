using MySql.Data.MySqlClient;
using System.Configuration;
using System.Data;

namespace new_nipt_processor.DAL
{
	public class ReservationsDAL
	{
		public static DataTable FilterReservations(string commandString, string dateStart, string dateEnd)
		{
			DataTable reservations = new DataTable();

			MySqlConnection mySqlConnection = new MySqlConnection(ConfigurationManager.AppSettings["myConnectionString"].ToString());
			mySqlConnection.Open();

			MySqlCommand cmd = new MySqlCommand(commandString, mySqlConnection);
			cmd.CommandType = CommandType.StoredProcedure;
			cmd.Parameters.AddWithValue("@inDateStart", dateStart);
			cmd.Parameters.AddWithValue("@inDateEnd", dateEnd);

			MySqlDataAdapter adt = new MySqlDataAdapter(cmd);
			adt.Fill(reservations);

			adt.Dispose();
			cmd.Dispose();
			mySqlConnection.Close();
			mySqlConnection.Dispose();

			return reservations;
		}

		public static DataTable FilterReservationsAll(string commandString)
		{
			DataTable reservations = new DataTable();

			MySqlConnection mySqlConnection = new MySqlConnection(ConfigurationManager.AppSettings["myConnectionString"].ToString());
			mySqlConnection.Open();

			MySqlCommand cmd = new MySqlCommand(commandString, mySqlConnection);
			cmd.CommandType = CommandType.StoredProcedure;

			MySqlDataAdapter adt = new MySqlDataAdapter(cmd);
			adt.Fill(reservations);

			adt.Dispose();
			cmd.Dispose();
			mySqlConnection.Close();
			mySqlConnection.Dispose();

			return reservations;
		}
	}
}
