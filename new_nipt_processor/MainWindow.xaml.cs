using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using new_nipt_processor.BAL;
using new_nipt_processor.Helper;
using Squirrel;
using System;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using Application = Microsoft.Office.Interop.Excel.Application;
using DataTable = System.Data.DataTable;
using Window = System.Windows.Window;

namespace new_nipt_processor
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		private string excelFileLocation;
		private string excelFilename;

		DataTable excelReservations;
		DataTable tableReservations;

		public MainWindow()
		{
			InitializeComponent();

			Task.Run(() => CheckAndApplyUpdate()).GetAwaiter().GetResult();
			this.Title += $" v.{GetVersionHelper.GetVersion()}";
		}

		private async Task CheckAndApplyUpdate()
		{
			try
			{
				bool updated = false;
				using (var updateManager = new UpdateManager(ConfigurationManager.AppSettings["fileServerLocation"].ToString()))
				{
					var updateInfo = await updateManager.CheckForUpdate();
					if (updateInfo.ReleasesToApply != null &&
						updateInfo.ReleasesToApply.Count > 0)
					{
						var releaseEntry = await updateManager.UpdateApp();
						updated = true;
					}
				}
				if (updated)
				{
					UpdateManager.RestartApp("pcr_processor.exe");
				}
			}
			catch
			{
			}
		}

		private void Process()
		{
			try
			{
				GetReservationsExcel();

				if (ChkOldDatabase.IsChecked == true)
				{
					GetReservationsTableAll();
				}
				else
				{
					GetReservationsTable();
				}

				//Check if there's duplicates in excel
				try
				{
					DataTable excelDuplicates = new DataTable();
					excelDuplicates = excelReservations.AsEnumerable()
														.GroupBy(r => new
														{
															illumina_report_id = r["illumina_report_id"]
														})
														.Where(g => g.Count() > 1)
														.Select(g => g.OrderBy(row => row["illumina_report_id"]).First())
														.CopyToDataTable();

					if (excelDuplicates.Rows.Count > 0)
					{
						MessageBox.Show("There is/are duplicate records in excel file.", "CREATE FILE",
							MessageBoxButton.OK, MessageBoxImage.Information);

						CreateExcelFileDuplicatesInDataTable(excelDuplicates, "DuplicatesInExcel");
						return;
					}
				}
				catch
				{
				}

				//Check if theres duplicates in reservations table
				try
				{
					DataTable tableDuplicates = new DataTable();
					tableDuplicates = tableReservations.AsEnumerable()
														.GroupBy(r => new
														{
															illumina_report_id = r["illumina_report_id"]
														})
														.Where(g => g.Count() > 1)
														.Select(g => g.OrderBy(row => row["illumina_report_id"]).First())
														.CopyToDataTable();

					if (tableDuplicates.Rows.Count > 0)
					{
						MessageBox.Show("There is/are duplicate records in reservations table.", "CREATE FILE",
							MessageBoxButton.OK, MessageBoxImage.Information);

						CreateExcelFileDuplicatesInDataTable(tableDuplicates, "DuplicatesInReservationsTable");
						return;
					}
				}
				catch
				{
				}

				foreach (DataRow row in excelReservations.Select())
				{
					foreach (DataRow row2 in tableReservations.Select())
					{
						try
						{
							if (row["illumina_report_id"].ToString().Equals(row2["illumina_report_id"].ToString()))
							{
								excelReservations.Rows.Remove(row);
								excelReservations.AcceptChanges();
								break;
							}
						}
						catch
						{
						}
					}
				}

				if (excelReservations.Rows.Count > 0)
				{
					CreateExcelFile(excelReservations);
				}

				MessageBox.Show("Done!", "CHECK FINISHED",
					MessageBoxButton.OK, MessageBoxImage.Information);
			}
			catch (Exception ex)
			{
				WriteLogFileHelper.WriteLogFile("Error: " + ex.Message.ToString());
			}
		}

		private void GetReservationsExcel()
		{
			excelReservations = new DataTable();

			string excelConnection = @"Provider=Microsoft.ACE.OLEDB.12.0; Data Source={0}; Extended Properties=Excel 12.0;";
			excelConnection = String.Format(excelConnection, excelFileLocation);

			using (OleDbConnection connection = new OleDbConnection(excelConnection))
			{
				connection.Open();

				OleDbCommand command = new OleDbCommand("SELECT * From [Sheet1$]", connection);
				OleDbDataAdapter adapter = new OleDbDataAdapter(command);
				adapter.Fill(excelReservations);
			}
		}

		private void GetReservationsTable()
		{
			DateTime ds = Convert.ToDateTime(DtpDateStart.Text.ToString());
			DateTime de = Convert.ToDateTime(DtpDateEnd.Text.ToString());

			tableReservations = ReservationsBAL.FilterUsers(ds.ToString("yyyy-MM-dd"), de.ToString("yyyy-MM-dd"));
		}

		private void GetReservationsTableAll()
		{
			tableReservations = ReservationsBAL.FilterReservationsAll();
		}

		private void CreateExcelFileDuplicatesInDataTable(DataTable content, string fileName)
		{
			SaveFileDialog saveFileDialog = new SaveFileDialog();
			saveFileDialog.FileName = fileName;
			saveFileDialog.DefaultExt = ".xlsx";
			saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";

			Nullable<bool> result = saveFileDialog.ShowDialog();
			if (result == true)
			{
				Application app = new Application();
				Workbook wb = app.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
				Worksheet ws = wb.Worksheets[1];
				ws.Range["A1"].Value = "illumina_report_id";

				if (content.Rows.Count > 0)
				{
					int cnt = 1;
					foreach (DataRow employee in content.Rows)
					{
						ws.Range["A" + (cnt + 1).ToString()].Value = employee[0].ToString();
						cnt++;
					}
				}
				wb.SaveAs(saveFileDialog.FileName, XlFileFormat.xlWorkbookDefault, Type.Missing,
					Type.Missing, false, false, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
				wb.Close();

				MessageBox.Show("Successfully created on file path " + saveFileDialog.FileName, "CREATE FILE",
					MessageBoxButton.OK, MessageBoxImage.Information);
			}
		}

		private void CreateExcelFile(DataTable content)
		{
			SaveFileDialog saveFileDialog = new SaveFileDialog();
			saveFileDialog.FileName = "NotExist";
			saveFileDialog.DefaultExt = ".xlsx";
			saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";

			Nullable<bool> result = saveFileDialog.ShowDialog();
			if (result == true)
			{
				Application app = new Application();
				Workbook wb = app.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
				Worksheet ws = wb.Worksheets[1];
				ws.Range["A1"].Value = "illumina_report_id";
				ws.Range["B1"].Value = "reserve_datetime";
				ws.Range["C1"].Value = "name";
				ws.Range["D1"].Value = "email";
				ws.Range["E1"].Value = "created_at";

				if (content.Rows.Count > 0)
				{
					int cnt = 1;
					foreach (DataRow employee in content.Rows)
					{
						ws.Range["A" + (cnt + 1).ToString()].Value = employee[0].ToString();
						ws.Range["B" + (cnt + 1).ToString()].Value = employee[1].ToString();
						ws.Range["C" + (cnt + 1).ToString()].Value = employee[2].ToString();
						ws.Range["D" + (cnt + 1).ToString()].Value = employee[3].ToString();
						ws.Range["E" + (cnt + 1).ToString()].Value = employee[4].ToString();
						cnt++;
					}
				}
				wb.SaveAs(saveFileDialog.FileName, XlFileFormat.xlWorkbookDefault, Type.Missing,
					Type.Missing, false, false, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
				wb.Close();

				MessageBox.Show("Successfully created on file path " + saveFileDialog.FileName, "CREATE FILE",
					MessageBoxButton.OK, MessageBoxImage.Information);
			}
		}

		private void BtnOpen_Click(object sender, RoutedEventArgs e)
		{
			OpenFileDialog openFileDialog = new OpenFileDialog();

			openFileDialog.DefaultExt = ".xlsx";
			openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";

			Nullable<bool> result = openFileDialog.ShowDialog();
			if (result == true)
			{
				string filename = openFileDialog.FileName;
				TxtFilePath.Text = filename;
				excelFileLocation = filename;
				excelFilename = filename.Split('\\')[filename.Split('\\').Length - 1];
			}
		}

		private void BtnProcess_Click(object sender, RoutedEventArgs e)
		{
			MessageBoxResult result = MessageBox.Show("Start processing?", "PROCESS",
				MessageBoxButton.YesNo, MessageBoxImage.Question);
			if (result == MessageBoxResult.Yes)
			{
				Process();
			}
		}

		private void BtnClose_Click(object sender, RoutedEventArgs e)
		{
			MessageBoxResult result = MessageBox.Show("Close this application?", "CLOSE",
				MessageBoxButton.YesNo, MessageBoxImage.Question);
			if (result == MessageBoxResult.Yes)
			{
				System.Windows.Application.Current.Shutdown();
			}
		}
	}
}
