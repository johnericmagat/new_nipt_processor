//using new_nipt_processor.BAL;
using new_nipt_processor.Helper;
//using new_nipt_processor.Model;
using Squirrel;
using System;
using System.Configuration;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

namespace new_nipt_processor
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
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
				
			}
			catch (Exception ex)
			{
				WriteLogFileHelper.WriteLogFile("Error: " + ex.Message.ToString());
			}
		}

		private void BtnOpen_Click(object sender, RoutedEventArgs e)
		{

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
				Application.Current.Shutdown();
			}
		}
	}
}
