using System;
using System.Diagnostics;
using System.Windows.Forms;
using NGWallpaper.Properties;
using System.Runtime.InteropServices;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Net;
using Microsoft.Win32;

namespace NGWallpaper
{
	/// <summary>
	/// 
	/// </summary>
	class ContextMenus
	{
        // Registry Parameters
        private const int SPI_SETDESKWALLPAPER = 20;
        private const int SPIF_UPDATEINIFILE = 0x01;
        private const int SPIF_SENDWININICHANGE = 0x02;

        /// <summary>
        /// HTTP Socket
        /// </summary>
        private static HttpClient Client = new HttpClient();

        /// <summary>
        /// Directory the program is in
        /// </summary>
        private static readonly string working_dir = System.IO.Directory.GetCurrentDirectory();
        
        /// <summary>
        /// Timer
        /// </summary>
        private Timer serviceTimer;

        /// <summary>
        /// Service flag
        /// </summary>
        bool isStarted = false;

        /// <summary>
        /// Creates this instance.
        /// </summary>
        /// <returns>ContextMenuStrip</returns>
        public ContextMenuStrip Create()
		{
			// Add the default menu options.
			ContextMenuStrip menu = new ContextMenuStrip();
			ToolStripMenuItem item;
			ToolStripSeparator sep;

            // Remove Icon margin
            menu.ShowImageMargin = false;

            // Folder
            item = new ToolStripMenuItem();
			item.Text = "Open Folder";
			item.Click += new EventHandler(Folder_click);
			menu.Items.Add(item);

			// Start
			item = new ToolStripMenuItem();
			item.Text = "Start";
			item.Click += new EventHandler(Start_click);
			menu.Items.Add(item);

			// Separator
			sep = new ToolStripSeparator();
			menu.Items.Add(sep);

			// Exit
			item = new ToolStripMenuItem();
			item.Text = "Exit";
			item.Click += new System.EventHandler(Exit_Click);
			menu.Items.Add(item);

            // Timer
            serviceTimer = new Timer();
            serviceTimer.Interval = 86400000;
            serviceTimer.Tick += OnTimedEventAsync;
            serviceTimer.Start();

            // Starts service
            Start_click(menu.Items[1], null);
            OnTimedEventAsync(null, null);

            return menu;
		}

        /// <summary>
        /// Handler for the timer event
        /// </summary>
        /// <param name="source">Source of event</param>
        /// <param name="e">Event</param>
        private async void OnTimedEventAsync(Object source, EventArgs e)
        {
            serviceTimer.Stop();
            string time = String.Format("{0:HH:mm:ss.fff}", System.DateTime.Now);
            string date = String.Format("{0:yyyy-MM-dd}", System.DateTime.Now);
            Console.WriteLine("The Timed event was raised at {0}", time);
            
            // Pop MessageBox
            DialogResult result = MessageBox.Show("Back up current wallpaper?", "NGWallpaper", MessageBoxButtons.YesNo);

            // Update image
            if (result == DialogResult.Yes)
            {
                // Read wallpaper path from registry
                string wallpaper_registry = (string)Registry.GetValue(@"HKEY_CURRENT_USER\Control Panel\Desktop", "WallPaper", working_dir + "/today.jpg");
                // today.jpg at current directory (Not in use)
                string source_file = working_dir + "/today.jpg";
                string dest_file = String.Format("{0}/{1}.jpg", working_dir, date);
                File.Move(wallpaper_registry, dest_file);
            }

            // Http
            await HTTP_Task();

            // Set Wallpaper
            SystemParametersInfo(SPI_SETDESKWALLPAPER,
                0,
                working_dir + "/today.jpg",
                SPIF_UPDATEINIFILE | SPIF_SENDWININICHANGE);

            MessageBox.Show("Complete", "NGWallpaper", MessageBoxButtons.OK);
            serviceTimer.Start();
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);

        private async Task HTTP_Task()
        {
            
            // Call asynchronous network methods in a try/catch block to handle exceptions
            try
            {
                string responseBody = await Client.GetStringAsync(Resources.POD_URL);
                //Regex pattern = new Regex(@"<meta property=""og: image"" content=""([\w\.?=%&=\-@/$,])"">");
                //string responseBody = @"<meta property=""og:image"" content=""https://yourshot.nationalgeographic.com/u/fQYSUbVfts-T7odkrFJckdiFeHvab0GWOfzhj7tYdC0uglagsDNd_d1hlDFbzHeE79tfZUXUfs6RF2z-MnlLI5uVLhShYScfCiC0jNl8OvvRnnTLH2qbz8iGifeG54JuUBffGw_vyQRd_Gim8LJeUEIxjeOqi_8d4pPKscZB40OrA8vMadCRfOu7KtoywnEittW8LrwRk9bnO1lnutFCDz4bbilSh8Rupg/""/>";
                Regex pattern = new Regex(@"og:image"" content=""([\w\.?=%&=\-@/$,:]+)");
                Match match = pattern.Match(responseBody);
                if (match.Groups.Count > 1)
                {
                    Console.WriteLine(match.Groups[1]);
                    using (WebClient web_client = new WebClient())
                    {
                        Uri uri = new Uri(match.Groups[1].ToString());
                        web_client.DownloadFileAsync(uri, working_dir + "/today.jpg");
                        while (web_client.IsBusy) { }
                    }
                }
                    
            }
            catch (HttpRequestException e)
            {
                Console.WriteLine("\nException Caught!");
                Console.WriteLine("Message :{0} ", e.Message);
                DialogResult result = MessageBox.Show("HTTP request failed. Do you want to retry?", "Error", MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                    OnTimedEventAsync(null, null);
                }

            }
        }

        /// <summary>
        /// Handles the Click event of the Open Folder control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        void Folder_click(object sender, EventArgs e)
		{
            Process.Start("explorer", working_dir);
		}

		/// <summary>
		/// Handles the Click event of the Start control.
		/// </summary>
		/// <param name="sender">The source of the event.</param>
		/// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
		void Start_click(object sender, EventArgs e)
		{
            ToolStripItem item = (ToolStripItem)sender;
            // Flip the flag
            isStarted = !isStarted;

            // Change Text
            if (isStarted)
            {
                item.Text = "Stop";
            }
            else
            {
                item.Text = "Start";
            }
            // Change the timer status
            serviceTimer.Enabled = isStarted;
		}

		/// <summary>
		/// Processes a menu item.
		/// </summary>
		/// <param name="sender">The sender.</param>
		/// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
		void Exit_Click(object sender, EventArgs e)
		{
            // Quit without further ado.
            serviceTimer.Dispose();
            Client.Dispose();
			Application.Exit();
		}
	}
}