/*
    NasaTvScheduleImport.  This program reads the NASA TV Schedule in Excel
    Format for the Space Shuttle and transfers the entries into Microsoft
    Outlook Calendar as Appointment items.

    Copyright (C) 2007  Ralph M. Hightower, Jr (Permanent Vacations)

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.

    mailto:Ralph.Hightower@gmail.com
 * Ralph Hightower
 * Chapin, SC 29036
*/
/* Change History
 * 20071227 - Ralph Hightower
 *      Added SuspendTimer(), ResumeTimer(); added appropriate calls to suspend and resume timing for user interactions
 * 20071218 - Ralph Hightower
 *      Added New Schedule Update button
 * 20071217 - Ralph Hightower
 *      Moved Outlook filter strings from code to resource
 * 20081110 - Ralph Hightower
 *      Fixed ImportMultipleSchedules to process last file
 * 20081127 - Ralph Hightower
 *      Added Mission Type (ISS, Hubble): ISS has docking
 */
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Resources;
using System.Runtime.InteropServices;
using System.Security;
using System.Security.Permissions;
using System.Text.RegularExpressions;
using System.Text;
using System.Windows.Forms;
using InteropOutlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Msdn.BclTeam;
using PermanentVacations.Nasa.Sts.Schedule;

[assembly: CLSCompliant(true)]
[assembly: PermissionSet(SecurityAction.RequestMinimum, Name="FullTrust")]
namespace PermanentVacations.Nasa.Sts.OutlookCalendar
{
	/// <summary>
	/// Code to handle the events for NASA Space Shuttle TV Schedule Transfer to Outlook Calendar
	/// </summary>
	public partial class MainForm : Form
	{
        private const string CRLF = "\r\n";

		/// <summary>
		/// True when application is shutting down
		/// </summary>
		bool applicationExiting = false;
		/// <summary>
		/// Stopwatch for timing lemgthy events (LoadExcelSchedule, TransferExcelToOutlook, RemoveSelectedOutlook,
		///		BulkImport)
		/// </summary>
		private Stopwatch stopWatch;
		/// <summary>
		/// State variable for if Outlook controls are being initialized.
		/// Prevents loading Outlook Schedule twice based on DateTimePicker being changed to today's date
		/// InitializeControls will finally load the Outlook Schedule
		/// </summary>
		private bool outlookControlsInitializing = true;
		/// <summary>
		/// Date of dtpOutlook
		/// </summary>
		private DateTime dtOutlookCalendar;

#if false   //  {
        /// <summary>
        /// Locations is SITE column in NASA STS TV Schedule DataGrid that should not be changed
        /// </summary>
        private string [] FixedLocations = {
            Properties.Resources.NASA_HQ,
            Properties.Resources.NASA_ISS_KU ,
            Properties.Resources.NASA_JPL,
            Properties.Resources.NASA_JSC,
            Properties.Resources.NASA_JSC_FG,
            Properties.Resources.NASA_KSC,
            Properties.Resources.NASA_TDRE,
            Properties.Resources.NASA_TDRW,
            Properties.Resources.NASA_TDRZ,
            Properties.Resources.NASA_VAFB
        };
#endif  //  }

        /// <summary>
        /// 
        /// </summary>
        private enum missionType
        {
            nonInitialized,
            missionISS,
            missionHubble
        };

        /// <summary>
        /// 
        /// </summary>
        private missionType typeMission;

		/// <summary>
		/// Getter/Setter of value containing dtpOutlook. Value
		/// </summary>
		private DateTime OutlookCalendar
		{
			get { return (dtOutlookCalendar); }
			set
			{
				if (value != null)
				{
					dtOutlookCalendar = value;
				}
			}
		}
		/// <summary>
		/// Checks the value of the last change from the DateTimePicker dtpOutlookCalendar
		/// and returns true if the value changed and updates the holding value with the new selection
		/// </summary>
		private bool CalendarChanged
		{
			get
			{
				bool calendarChanged = (OutlookCalendar != dtpOutlook.Value);
				if (calendarChanged)
				{
					OutlookCalendar = dtpOutlook.Value;
				}
				return (calendarChanged);
			}
		}

		/// <summary>
		/// Updates the progress bar with the current operation
		/// When set to 0, it resets the progress bar
		/// if value = -1, sets the progress bar to 0
		/// </summary>
		private int ProgressBar
		{
			get { return (toolStripProgressBar.Value); }
			set
			{
				if ((value >= toolStripProgressBar.Minimum) && (value <= toolStripProgressBar.Maximum))
				{
					toolStripProgressBar.Value = value;
					toolStripProgressBar.PerformStep();
				}
				else
					SetupProgressBar(0);
			}
		}

		/// <summary>
		/// Getter/Setter for StatusBar
		/// Used to primarily display messages on the status bar
		/// </summary>
		private string Status
		{
			get { return (toolStripStatusLabel.Text); }
			set
			{
				toolStripStatusLabel.Text = value;
				statusStrip.Refresh();
			}
		}

		/// <summary>
		/// Creates and Windows initialization of the  controls
		/// </summary>
		public MainForm()
		{
			InitializeComponent();
		}

		/// <summary>
		/// Handler for Form Load
		/// Initializes the Excel and Outlook controls
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void MainForm_Load(object sender, EventArgs e)
		{
			InitializeControls();
		}

		/// <summary>
		/// Initializes the controls
		/// </summary>
		private void InitializeControls()
		{
			Busy(Properties.Resources.ID_BUSY);

			DisableMenus();

			InitializeExcelControls();
			LoadOutlookControls();

#if false   //  {
            //  If it's May or June in 2009, then it's probably STS-125, the Hubble Servicing Mission
            DateTime dtToday = DateTime.Now;
            if ((dtToday.Year == 2009) && ((dtToday.Month == 5) || (dtToday.Month == 6)))
                rbHubble.Checked = true;
            else
                rbISS.Checked = true;
#endif  //  }

			Ready();
		}

		/// <summary>
		/// Handler for the Help About menu
		/// Displays the About Box
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
		{
			AboutBox dlgAbout = new AboutBox();
			dlgAbout.ShowDialog();
			dlgAbout = null;
		}

		/// <summary>
		/// Handler for Exit Application Button
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnExitApplication_Click(object sender, EventArgs e)
		{
			ReleaseResources();
			SaveUserSettings();
			System.Windows.Forms.Application.Exit();
		}

		/// <summary>
		/// Tool tip for the Exit Application button
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnExitApplication_MouseHover(object sender, EventArgs e)
		{
			Status = Properties.Resources.MOUSEHOVER_EXITAPPLICATION;
		}

		/// <summary>
		/// Handler for the MouseLeave (Clears Tool tip) Exit Application
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnExitApplication_MouseLeave(object sender, EventArgs e)
		{
			Status = "";
		}

		/// <summary>
		/// Sets cursor to busy and status bar to busy
		/// </summary>
		private void Busy(string busy)
		{
			Cursor = Cursors.WaitCursor;
			Status = busy;
		}

		/// <summary>
		/// Sets cursor to ready and status bar to ready
		/// </summary>
		private void Ready()
		{
			Cursor = Cursors.Default;
			Status = Properties.Resources.ID_READY;
		}

		/// <summary>
		/// Allocates a stopwatch and starts timing
		/// </summary>
		private void StartTimer()
		{
			if (stopWatch == null)
			{
				stopWatch = new Stopwatch();
			}
			if (stopWatch.IsRunning)
			{
				stopWatch.Stop();
				stopWatch.Reset();
			}
			stopWatch.Start();
		}

		private void SuspendTimer()
		{
			if (stopWatch != null)
			{
				if (stopWatch.IsRunning)
					stopWatch.Stop();
			}
		}

		private void ResumeTimer()
		{
			if (stopWatch != null)
			{
				if (!stopWatch.IsRunning)
					stopWatch.Start();
			}
		}

		/// <summary>
		/// Stops timing, displays elapsed time on status bar and nulls the stopwatch
		/// </summary>
		private void StopTimer(string elapsedTime)
		{
			if (stopWatch != null)
			{
				stopWatch.Stop();
				TimeSpan tsElapsed = stopWatch.Elapsed;
				Status = elapsedTime + tsElapsed.ToString();
			}
			stopWatch = null;
		}

		/// <summary>
		/// Disables menus that are not implemented
		/// </summary>
		private void DisableMenus()
		{
			newToolStripMenuItem.Enabled = false;
			saveToolStripMenuItem.Enabled = false;
			saveAsToolStripMenuItem.Enabled = false;
			printToolStripMenuItem.Enabled = false;
			printPreviewToolStripMenuItem.Enabled = false;
			undoToolStripMenuItem.Enabled = false;
			redoToolStripMenuItem.Enabled = false;
			cutToolStripMenuItem.Enabled = false;
			copyToolStripMenuItem.Enabled = false;
			pasteToolStripMenuItem.Enabled = false;
			selectAllToolStripMenuItem.Enabled = false;
			customizeToolStripMenuItem.Enabled = false;
			optionsToolStripMenuItem.Enabled = false;
			contentsToolStripMenuItem.Enabled = false;
			indexToolStripMenuItem.Enabled = false;
			searchToolStripMenuItem.Enabled = false;
			selectAllToolStripMenuItem.Enabled = false;
		}

		/// <summary>
		/// Free up resources used
		/// Force Garbage Collection to have spawned Excel processes exit
		/// </summary>
		private void ReleaseResources()
		{
			applicationExiting = true;
			statusStrip.Refresh();
			GC.Collect();
		}

		/// <summary>
		/// Saves User's Preferences
		/// </summary>
		private void SaveUserSettings()
		{
			Properties.Settings.Default.Save();
		}

		/// <summary>
		/// Initializes the Progress Status Bar
		/// </summary>
		/// <param name="maxValue"></param>
		private void SetupProgressBar(int maxValue)
		{
			//	Gets or sets the lower bound of the range that is defined for this ToolStripProgressBar.
			toolStripProgressBar.Minimum = 0;
			//	Gets or sets the upper bound of the range that is defined for this ToolStripProgressBar.
			toolStripProgressBar.Maximum = maxValue;
			//	Gets or sets the amount by which to increment the current value of the ToolStripProgressBar
			//	when the PerformStep method is called.
			toolStripProgressBar.Step = 1;
			ProgressBar = 0;
		}

		#region Excel Control Functions
		/// <summary>
		/// Handler for Open Nasa Sts TV Schedule button
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnOpenNasaTvSchedule_Click(object sender, EventArgs e)
		{
            Busy(Properties.Resources.ID_BUSY_READING_EXCEL);
            StartTimer();
            OpenNasaTvSchedule();
            Ready();
            StopTimer(Properties.Resources.TIMER_ELAPSED_TIME_READING_EXCEL);
		}

		/// <summary>
		/// Handler for the File Open menu item
		/// Open a NASA TV Schedule and process it
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void openToolStripMenuItem_Click(object sender, EventArgs e)
		{
            Busy(Properties.Resources.ID_BUSY_READING_EXCEL);
            StartTimer();
            OpenNasaTvSchedule();
            Ready();
            StopTimer(Properties.Resources.TIMER_ELAPSED_TIME_READING_EXCEL);
        }

		/// <summary>
		/// Tool tip for the Open NASA TV Schedule button
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnOpenNasaTvSchedule_MouseHover(object sender, EventArgs e)
		{
			Status = Properties.Resources.MOUSEHOVER_OPENNASATVSCHEDULE;
		}

		/// <summary>
		/// Handler for the MouseLeave (Clears Tool tip) Open NASA TV Schedule file
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnOpenNasaTvSchedule_MouseLeave(object sender, EventArgs e)
		{
			Status = "";
		}

		/// <summary>
		/// Selects all items in the Nasa Sts TV Schedule Data Grid
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnSelectAllExcel_Click(object sender, EventArgs e)
		{
			statusStrip.ResetText();

			Busy(Properties.Resources.ID_BUSY);
			SelectAllExcel();
			Ready();
		}

		/// <summary>
		/// Tool tip for the Select All Excel button
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnSelectAllExcel_MouseHover(object sender, EventArgs e)
		{
			Status = Properties.Resources.MOUSEHOVER_SELECTALLEXCEL;
		}

		/// <summary>
		/// Handler for the MouseLeave (Clears Tool tip) Select All Excel entries
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnSelectAllExcel_MouseLeave(object sender, EventArgs e)
		{
			Status = "";
		}

		/// <summary>
		/// Unselect all entries in Excel Schedule
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnUnselectAllExcel_Click(object sender, EventArgs e)
		{
			statusStrip.ResetText();

			Busy(Properties.Resources.ID_BUSY);
			UnselectAllExcel();
			Ready();
		}

		/// <summary>
		/// Tool tip for the Unselect All Excel button
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnUnselectAllExcel_MouseHover(object sender, EventArgs e)
		{
			Status = Properties.Resources.MOUSEHOVER_UNSELECTALLEXCEL;
		}

		/// <summary>
		/// Handler for the MouseLeave (Clears Tool tip) Unselect All Excel entries
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnUnselectAllExcel_MouseLeave(object sender, EventArgs e)
		{
			Status = "";
		}

		/// <summary>
		/// Handler for the MouseLeave (Clears Tool tip) Bulk Import
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnBulkImport_MouseLeave(object sender, EventArgs e)
		{
			Status = "";
		}

        /// <summary>
        /// Updates the Outlook Calendar with the new Excel schedule loaded
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnNewScheduleUpdate_Click(object sender, EventArgs e)
        {
            UpdateNewSchedule();
        }

        /// <summary>
        /// Clears hover tip
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnNewScheduleUpdate_Leave(object sender, EventArgs e)
        {
            Status = "";
        }

        /// <summary>
        /// Hover tip for New Schedule Update
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnNewScheduleUpdate_MouseHover(object sender, EventArgs e)
        {
            Status = Properties.Resources.MOUSEHOVER_NEWSCHEDULEUPDATE;
        }

        /// <summary>
		/// Handler for Refresh Outlook Categories
		/// Reloads the categories from Outlook into the CheckListBox
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnRefreshOutlookCategories_Click(object sender, EventArgs e)
		{
			statusStrip.ResetText();

			Busy(Properties.Resources.ID_BUSY);

			LoadOutlookCategories();

			Ready();
		}

		/// <summary>
		/// Tool tip for the Refresh Outlook Categories button
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnRefreshOutlookCategories_MouseHover(object sender, EventArgs e)
		{
			Status = Properties.Resources.MOUSEHOVER_REFRESHCATEGORIES;
		}

		/// <summary>
		/// Handler for the MouseLeave (Clears Tool tip) Outlook Categories Refresh
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnRefreshOutlookCategories_MouseLeave(object sender, EventArgs e)
		{
			Status = "";
		}

		/// <summary>
		/// Handler to removed selected entries in the Outlook Schedule Data Grid from the Outlook Calendar
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnRemoveMarkedEntries_Click(object sender, EventArgs e)
		{
			StartTimer();
			Busy(Properties.Resources.ID_BUSY_REMOVING_APPOINTMENTS);

			RemoveOutlookEntries();

			LoadOutlookSchedule();

			Ready();
			StopTimer(Properties.Resources.TIMER_ELAPSED_TIME_REMOVING_SCHEDULE);
		}

		/// <summary>
		/// Tool tip for the Deleted Selected Appointment Items
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnRemoveMarkedEntries_MouseHover(object sender, EventArgs e)
		{
			Status = Properties.Resources.MOUSEHOVER_REMOVESELECTEDITEMS;
		}

		/// <summary>
		/// Handler for the MouseLeave (Clears Tool tip) Remove Marked Entries
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnRemoveMarkedEntries_MouseLeave(object sender, EventArgs e)
		{
			Status = "";
		}

		/// <summary>
		/// Selects all items in the Outlook Data Grid
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnSelectAllOutlook_Click(object sender, EventArgs e)
		{
			statusStrip.ResetText();

			Busy(Properties.Resources.ID_BUSY);
			SelectAllOutlook();
			Ready();
		}

		/// <summary>
		/// Tool tip for the Select All Outlook button
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnSelectAllOutlook_MouseHover(object sender, EventArgs e)
		{
			Status = Properties.Resources.MOUSEHOVER_SELECTALLOUTLOOK;
		}

		/// <summary>
		/// Handler for the MouseLeave (Clears Tool tip) Select All Outlook Entries
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnSelectAllOutlook_MouseLeave(object sender, EventArgs e)
		{
			Status = "";
		}

		/// <summary>
		/// Handler for the Smart Select button
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnSmartSelect_Click(object sender, EventArgs e)
		{
			if (dgvExcelSchedule.RowCount > 0)
			{
				Busy(Properties.Resources.ID_BUSY);
				SmartSelect();
				Ready();
			}
			else
			{
				SuspendTimer();
				MessageBox.Show(Properties.Resources.ERR_NO_EXCEL_SCHEDULE,
					Properties.Resources.ERR_MESSAGE_BOX_HDR_ERROR, MessageBoxButtons.OK, MessageBoxIcon.Error,
					MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
				ResumeTimer();
			}
		}

		/// <summary>
		/// Tool tip for the Smart Select button
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnSmartSelect_MouseHover(object sender, EventArgs e)
		{
			Status = Properties.Resources.MOUSEHOVER_SMARTSELECT;
		}

		/// <summary>
		/// Handler for the MouseLeave (Clears Tool tip) Smart Select
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnSmartSelect_MouseLeave(object sender, EventArgs e)
		{
			Status = "";
		}

		/// <summary>
		/// Handler to transfer the selected items from the Nasa Sts TV Schedule to the Outlook Calendar
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnTransferTVSchedule_Click(object sender, EventArgs e)
		{
			statusStrip.ResetText();

			StartTimer();
			Busy(Properties.Resources.ID_BUSY_TRANSFERRING_SCHEDULE);

			TransferExcelToOutlook();

			Ready();
			StopTimer(Properties.Resources.TIMER_ELAPSED_TIME_TRANFERRING);
		}

		/// <summary>
		/// Tool tip for the  for the Transfer TV Schedule button
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnTransferTVSchedule_MouseHover(object sender, EventArgs e)
		{
			Status = Properties.Resources.MOUSEHOVER_TRANSFERTVSCHEDULE;
		}

		/// <summary>
		/// Handler for the MouseLeave (Clears Tool tip) Transfer TV Schedule
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnTransferTVSchedule_MouseLeave(object sender, EventArgs e)
		{
			Status = "";
		}

		/// <summary>
		/// Unselect all entries in Outlook Data Grid
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnUnselectAllOutlook_Click(object sender, EventArgs e)
		{
			statusStrip.ResetText();

			Busy(Properties.Resources.ID_BUSY);
			UnselectAllOutlook();
			Ready();
		}

		/// <summary>
		/// Tool tip for the Select All Outlook button
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnUnselectAllOutlook_MouseHover(object sender, EventArgs e)
		{
			Status = Properties.Resources.MOUSEHOVER_UNSELECTALLOUTLOOK;
		}

		/// <summary>
		/// Handler for the MouseLeave (Clears Tool tip) Unselect All Outlook
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnUnselectAllOutlook_MouseLeave(object sender, EventArgs e)
		{
			Status = "";
		}

		/// <summary>
		/// Tool tip for the Docked checkbox
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void chkDocked_MouseHover(object sender, EventArgs e)
		{
			Status = Properties.Resources.MOUSEHOVER_DOCKED;
		}

		/// <summary>
		/// Handler for the MouseLeave (Clears Tool tip) Docked Status Checkbox
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void chkDocked_MouseLeave(object sender, EventArgs e)
		{
			Status = "";
		}

		/// <summary>
		/// Tool tip for the Orbit checkbox
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void chkOrbit_MouseHover(object sender, EventArgs e)
		{
			Status = Properties.Resources.MOUSEHOVER_ORBIT;
		}

		/// <summary>
		/// Handler for the MouseLeave (Clears Tool tip) In Orbit Checkbox
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void chkOrbit_MouseLeave(object sender, EventArgs e)
		{
			Status = "";
		}

        /// <summary>
        /// Handler of NASA STS TV Schedule DataGridView CellEnter: 
        ///     If the column is the location of the event, check to see if the location is not a fixed ground or space location:
        ///     NASA HQ, JSC, KSC, or space locations, such as, TDRS.
        ///     If the cell contents is ISS or STS, allow editing.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvExcelSchedule_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (dgvExcelSchedule[e.ColumnIndex, e.RowIndex].OwningColumn.HeaderText == SITE_TV.HeaderText)
            {
                string site = dgvExcelSchedule[e.ColumnIndex, e.RowIndex].Value.ToString();
                if ((site == Properties.Resources.NASA_ISS) || (site == Properties.Resources.NASA_STS))
                    dgvExcelSchedule.BeginEdit(true);
#if false   //  {
            bool exit = false;
                for (int idxFixedLocations = FixedLocations.GetLowerBound(0);
                    idxFixedLocations <= FixedLocations.GetUpperBound(0) && !exit;
                    idxFixedLocations++)
                {
                    if (dgvExcelSchedule[e.ColumnIndex, e.RowIndex].Value.ToString() == FixedLocations[idxFixedLocations])
                        exit = true;
                }
                if (!exit)
                    dgvExcelSchedule.BeginEdit(true);
#endif
            }
        }

        /// <summary>
        /// Handler for NASA STS TV Schedule DataGridView CellLeave:
        ///     If the current cell is in edit mode, then commit the changes
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvExcelSchedule_CellLeave(object sender, DataGridViewCellEventArgs e)
        {
            if (dgvExcelSchedule.CurrentCell.IsInEditMode)
                dgvExcelSchedule.EndEdit(DataGridViewDataErrorContexts.Commit);
        }

        /// <summary>
		/// Handler for ColumnSortModeChanged
		/// May need to develop this for DateTime comparison if string compare is used instead of DateTime
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void dgvExcelSchedule_ColumnSortModeChanged(object sender, DataGridViewColumnEventArgs e)
		{
			DataGridViewColumnSortMode sortMode = e.Column.SortMode;
			switch (sortMode)
			{
				case DataGridViewColumnSortMode.Automatic:
					break;
				case DataGridViewColumnSortMode.NotSortable:
					break;
				case DataGridViewColumnSortMode.Programmatic:
					break;
			}
		}

		/// <summary>
		/// Handle changes in the data grid for the Nasa Sts TV Schedule
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void dgvExcelSchedule_CurrentCellDirtyStateChanged(object sender, EventArgs e)
		{
			//  Accept the change of the checkbox (ID_ADD)
			if (dgvExcelSchedule.IsCurrentCellDirty)
			{
				dgvExcelSchedule.CommitEdit(DataGridViewDataErrorContexts.Commit);
			}
		}

		/// <summary>
		/// This error handler is needed because there is something that generates errors when adding entries in the data grid
		///
		/// I need to figure out what the problem is
		/// Note 20071201:
		///		This error occurred early in the development process and did not occur reading the STS-122 tvsched_rev0.xls
		///		published 11/30/2007.
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void dgvExcelSchedule_DataError(object sender, DataGridViewDataErrorEventArgs e)
		{

		}

		/// <summary>
		/// Tool tip for the Excel Data Grid
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void dgvExcelSchedule_MouseHover(object sender, EventArgs e)
		{
			Status = Properties.Resources.MOUSEHOVER_EXCELDATAGRID;
		}

		/// <summary>
		/// Handler for the MouseLeave (Clears Tool tip) Excel Data Grid
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void dgvExcelSchedule_MouseLeave(object sender, EventArgs e)
		{
			Status = "";
		}

		/// <summary>
		/// Handle DataGridView sorting
		/// May need to develop this for DateTime comparison if string compare is used instead of DateTime
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void dgvExcelSchedule_SortCompare(object sender, DataGridViewSortCompareEventArgs e)
		{
			string header = dgvExcelSchedule.SortedColumn.HeaderText;
			if ((header == BEGIN_DATE_TV.Name) || (header == END_DATE_TV.Name))
			{
				e.SortResult = System.DateTime.Compare((DateTime)e.CellValue1, (DateTime)e.CellValue2);
			}
			else
			{
				e.SortResult = System.String.Compare(e.CellValue1.ToString(), e.CellValue2.ToString());
			}
			e.Handled = true;
		}

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void rbISS_CheckedChanged(object sender, EventArgs e)
        {
            typeMission = (rbISS.Checked ? missionType.missionISS : missionType.missionHubble);
        }

        /// <summary>
        /// Handler of Mouse Hover over Radio Button: ISS
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void rbISS_MouseHover(object sender, EventArgs e)
        {
            Status = Properties.Resources.MOUSEHOVER_MISSIONISS;
        }

        /// <summary>
        /// Handler for Mouse Hover over Radio Button: ISS
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void rbISS_MouseLeave(object sender, EventArgs e)
        {
            Status = "";
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void rbHubble_CheckedChanged(object sender, EventArgs e)
        {
            typeMission = (rbHubble.Checked ? missionType.missionHubble : missionType.missionISS);
        }

        /// <summary>
        /// Handler for Mouse Hover over Radio Button: Standalone/Hubble
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void rbHubble_MouseHover(object sender, EventArgs e)
        {
            Status = Properties.Resources.MOUSEHOVER_MISSIONSTANDALONE;
        }

        /// <summary>
        /// Handler for Mouse Hover over Radio Button: Standalone/Hubble
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void rbHubble_MouseLeave(object sender, EventArgs e)
        {
            Status = "";
        }

        /// <summary>
		/// Initializes the Data Grid for the schedule read from the Nasa Sts TV Schedule
		/// </summary>
		private void InitializeExcelControls()
		{
			lblNasaStsTVScheduleFile.Text = "";

			//  Default always to Microsoft.Office.Interop.Excel
			//  Reason: The Interop method opens an Excel file
			//  I have not found a way to open an Excel file using Microsoft.Office.Tools.Excel
			chkInteropExcel.Checked = true;

            if (rbISS.Checked)
                typeMission = missionType.missionISS;
            else
                typeMission = missionType.missionHubble;

			InitializeExcelDataGrid();
		}

		/// <summary>
		/// Clears the entries in the NASA STS TV Schedule Grid
		/// </summary>
		private void InitializeExcelDataGrid()
		{
			dgvExcelSchedule.Rows.Clear();
		}

        /// <summary>
        /// Shows Open File Dialog and loads new NASA TV Schedule
        /// </summary>
        /// <returns>true if user selects OK and loads schedule</returns>
        private bool OpenNasaTvSchedule()
        {
            bool openNasaTvSchedule = false;

            ofExcelSchedule.Multiselect = false;
            ofExcelSchedule.ReadOnlyChecked = true;
            ofExcelSchedule.Title = Properties.Resources.ID_OPEN_SINGLE_EXCEL_FILE;
            ofExcelSchedule.InitialDirectory = Properties.Settings.Default.MyDocuments;

            SuspendTimer();

            DialogResult drExcelSchedule = ofExcelSchedule.ShowDialog();

            ResumeTimer();

            if (drExcelSchedule == DialogResult.OK)
            {
                openNasaTvSchedule = true;

                string excelSchedule = ofExcelSchedule.FileName;

                FileInfo fiExcelSchedule = new FileInfo(excelSchedule);
                Properties.Settings.Default.MyDocuments = fiExcelSchedule.DirectoryName;
                fiExcelSchedule = null;

                LoadExcelSchedule(excelSchedule);
            }

            return (openNasaTvSchedule);
        }

        /// <summary>
        /// Select All entries in Excel Data Grid (NASA STS TV Schedule)
        /// </summary>
        private void SelectAllExcel()
        {
            for (int index = 0; index < dgvExcelSchedule.Rows.Count; index++)
            {
                dgvExcelSchedule.Rows[index].Cells[ID_ADD.Name].Value = true;
            }
        }

        /// <summary>
        /// Unselect all entries in Excel Data Grid (NASA STS TV Schedule)
        /// </summary>
        private void UnselectAllExcel()
        {
            for (int index = 0; index < dgvExcelSchedule.Rows.Count; index++)
            {
                dgvExcelSchedule.Rows[index].Cells[ID_ADD.Name].Value = false;
            }
        }

        /// <summary>
		/// Populates the data grid with data from the Nasa Sts TV Schedule class
		/// </summary>
		/// <param name="excelFile">NASA TV Schedule Excel spreadsheet</param>
		private void LoadExcelSchedule(string excelFile)
		{
			InitializeExcelDataGrid();
			NasaStsTVSchedule tvSchedule = null;

			FileInfo fiExcelSchedule = new FileInfo(excelFile);
			DirectoryInfo diExcelSchedule = fiExcelSchedule.Directory;
			string dirFile = fiExcelSchedule.Directory.Name + "/" + fiExcelSchedule.Name + " - ";
            lblNasaStsTVScheduleFile.Text = dirFile;
            string strTitle = dirFile + Properties.Resources.ID_PROGRAMTITLE;
			this.Text = strTitle;

			try
			{
				string viewingTimeZone;

				//  If no Time Zone has been selected, use the default
				if (cmbxTimeZones.SelectedIndex > -1)
				{
					viewingTimeZone = cmbxTimeZones.Items[cmbxTimeZones.SelectedIndex].ToString();
					Properties.Settings.Default.VewersTimeZone = viewingTimeZone;
				}
				else
				{
					viewingTimeZone = Properties.Resources.ID_TZ_DEFAULT;
				}

				//  Use the Nasa Sts TV Schedule class to populate the Nasa Sts TV Schedule Data Grid
				tvSchedule = new NasaStsTVSchedule(excelFile, viewingTimeZone, typeMission == missionType.missionISS);

				NasaStsTVScheduleEntry scheduleRow = tvSchedule.ReadScheduleRow();

				SetupProgressBar(tvSchedule.NumberRows);

				bool noErrors = true;
				if (scheduleRow != null)
					noErrors = scheduleRow.TypeEntry != ScheduleType.error;

				dgvExcelSchedule.SuspendLayout();

				while (!tvSchedule.EOF() && noErrors)
				{
					if (scheduleRow != null)
					{
						noErrors = (scheduleRow.TypeEntry != ScheduleType.error);
						if (noErrors)
						{
							dgvExcelSchedule.Rows.Add(false, scheduleRow.Orbit, scheduleRow.BeginDate,
								scheduleRow.EndDate, false, scheduleRow.Revised(), scheduleRow.Subject,
								scheduleRow.Site);
							chkDocked.Checked = tvSchedule.IsDocked();
							chkOrbit.Checked = tvSchedule.InSpace();
							ProgressBar = tvSchedule.CurrentEntry;
						}
					}
					if (noErrors)
						scheduleRow = tvSchedule.ReadScheduleRow();
				}
				if (!noErrors)
				{
					if (scheduleRow != null)
					{
						SuspendTimer();
						MessageBox.Show(scheduleRow.Subject, Properties.Resources.ERR_EXCEPTION, MessageBoxButtons.OK,
							MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, (MessageBoxOptions)0);
						ResumeTimer();
					}
				}

				ProgressBar = -1;
			}
			catch (COMException comExp)
			{
				SuspendTimer();
				MessageBox.Show(comExp.Message + comExp.StackTrace, Properties.Resources.ERR_COM_EXCEPTION,
					MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
					(MessageBoxOptions)0);
				ResumeTimer();
				if (Properties.Settings.Default.CopyExceptionsToClipboard)
                    Clipboard.SetText(comExp.Message + CRLF + comExp.StackTrace, TextDataFormat.Text);
			}
			finally
			{
				dgvExcelSchedule.ResumeLayout();
				dgvExcelSchedule.Refresh();
				tvSchedule.Close();
				tvSchedule = null;
			}
		}

		#endregion

		#region Outlook Functions
		/// <summary>
		/// Handler for the Bulk Import
		/// Imports multiple schedules at a time
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnBulkImport_Click(object sender, EventArgs e)
		{
			ofExcelSchedule.Multiselect = true;
			ofExcelSchedule.ReadOnlyChecked = true;
			ofExcelSchedule.Title = Properties.Resources.ID_OPEN_MULTIPLE_EXCEL_FILES;

			DialogResult drExcelSchedules = ofExcelSchedule.ShowDialog();
			if (drExcelSchedules == DialogResult.OK)
			{
				Busy(Properties.Resources.ID_THIS_WILL_TAKE_A_LONG_WHILE);
				StartTimer();
				string[] nasaTvSchedules = ofExcelSchedule.FileNames;
				Array.Sort(nasaTvSchedules);

				ImportMultipleSchedules(nasaTvSchedules);
				Ready();
				StopTimer(Properties.Resources.TIMER_ELAPSED_TIME_BULK_IMPORT);
			}
		}

		/// <summary>
		/// Tool tip for the for the Bulk Import button
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnBulkImport_MouseHover(object sender, EventArgs e)
		{
			Status = Properties.Resources.MOUSEHOVER_BULKIMPORT;
		}

		/// <summary>
		/// Tool tip for the Outlook Categories CheckedListBox
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void chklbOutlookCategories_MouseHover(object sender, EventArgs e)
		{
			Status = Properties.Resources.MOUSEHOVER_OUTLOOKCATEGORIES;
		}

		/// <summary>
		/// Handler for the MouseLeave (Clears Tool tip) Outlook Categores CheckedListBox
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void chklbOutlookCategories_MouseLeave(object sender, EventArgs e)
		{
			Status = "";
		}

		/// <summary>
		/// Handler for the Outlook Categories CheckListBox SelectedValueChanged
		/// Reload the Outlook Schedule with the new category selections
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void chklbOutlookCategories_SelectedValueChanged(object sender, EventArgs e)
		{
			Busy(Properties.Resources.ID_BUSY);
			LoadOutlookSchedule();
			Ready();
		}

		/// <summary>
		/// Tool tip for the Viewer's Time Zone
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void cmbxTimeZones_MouseHover(object sender, EventArgs e)
		{
			Status = Properties.Resources.MOUSEHOVER_VIEWINGTIMEZONE;
		}

		/// <summary>
		/// Handler for the MouseLeave (Clears Tool tip) Viewer's Time Zone
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void cmbxTimeZones_MouseLeave(object sender, EventArgs e)
		{
			Status = "";
		}

		/// <summary>
		/// Handler for the Viewer's Time Zone Selection Changed
		/// Saves the selection in the user's settings
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void cmbxTimeZones_SelectedIndexChanged(object sender, EventArgs e)
		{
			Properties.Settings.Default.VewersTimeZone = cmbxTimeZones.Items[cmbxTimeZones.SelectedIndex].ToString();
		}

		/// <summary>
		/// Handle changes in the data grid for the Outlook Schedule
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void dgvOutlook_CurrentCellDirtyStateChanged(object sender, EventArgs e)
		{
			//  Accept the change of the checkbox (REMOVE_OL, REMINDER_OL)
			if (dgvOutlook.IsCurrentCellDirty)
			{
				dgvOutlook.CommitEdit(DataGridViewDataErrorContexts.Commit);
			}
		}

		/// <summary>
		/// This is probably needed also to process the handling of DataErrors in the Outlook grid
		/// 
		/// Note 20071201:
		///		This was added early in the development process because of problems encountered with the Excel Data Grid.
		///		A breakpoint was not encountered loading the Outlook schedule for STS-122
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void dgvOutlook_DataError(object sender, DataGridViewDataErrorEventArgs e)
		{

		}

		/// <summary>
		/// Tool tip for the Outlook Data Grid
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void dgvOutlook_MouseHover(object sender, EventArgs e)
		{
			Status = Properties.Resources.MOUSEHOVER_OUTLOOKDATAGRID;
		}

		/// <summary>
		/// Handler for the MouseLeave (Clears Tool tip) Outlook Data Grid
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void dgvOutlook_MouseLeave(object sender, EventArgs e)
		{
			Status = "";
		}

		/// <summary>
		/// Handler for the Outlook Calendar CloseUp
		/// If the value changed, reloads the schedule from Outlook
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void dtpOutlook_CloseUp(object sender, EventArgs e)
		{
			if (CalendarChanged && !outlookControlsInitializing)
			{
				Busy(Properties.Resources.ID_BUSY);
				LoadOutlookSchedule();
				Ready();
			}
		}

		/// <summary>
		/// Handler for the Leave event for the Outlook DateTimePicker
		/// Check to see if the value changed and reload the schedule if necessary
		/// This should cover the situation for where the date is entered manually
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void dtpOutlook_Leave(object sender, EventArgs e)
		{
			//	This event could fire when the Exit button is pressed
			if (CalendarChanged && !outlookControlsInitializing && !applicationExiting)
			{
				Busy(Properties.Resources.ID_BUSY);
				LoadOutlookSchedule();
				Ready();
			}
		}

		/// <summary>
		/// Tool tip for the Mission Date
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void dtpOutlook_MouseHover(object sender, EventArgs e)
		{
			Status = String.Format(CultureInfo.InstalledUICulture, Properties.Resources.MOUSEHOVER_MISSIONDATE,
			Properties.Settings.Default.LookAheadWeeks);
		}

		/// <summary>
		/// Handler for the MouseLeave (Clears Tool tip) Mission Date
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void dtpOutlook_MouseLeave(object sender, EventArgs e)
		{
			Status = "";
		}

		/// <summary>
        /// <summary>
        /// Select all entries in Outlook Data Grid
        /// </summary>
        private void SelectAllOutlook()
        {
            for (int index = 0; index < dgvOutlook.Rows.Count; index++)
            {
                dgvOutlook.Rows[index].Cells[REMOVE_OL.Name].Value = true;
            }
        }

        /// <summary>
        /// Uses first Begin Date and last Begin Date in Excel DataGrid
        /// to select those entries in the Outlook DataGrid that are between the first and last dates
        /// </summary>
        private void SmartSelect()
        {
            if (dgvExcelSchedule.RowCount > 0)
            {
                DateTime dtBegin = (DateTime)dgvExcelSchedule[BEGIN_DATE_TV.Name, 0].Value;
                DateTime dtEnd = (DateTime)dgvExcelSchedule[BEGIN_DATE_TV.Name, dgvExcelSchedule.RowCount - 1].Value;

                int indexOutlook;
                int outlookEntries = dgvOutlook.RowCount;
                for (indexOutlook = 0; indexOutlook < outlookEntries; indexOutlook++)
                {
                    DateTime dtEntry = (DateTime)dgvOutlook[BEGIN_DATE_OL.Name, indexOutlook].Value;
                    if ((dtBegin <= dtEntry) && (dtEntry <= dtEnd))
                        dgvOutlook.Rows[indexOutlook].Cells[REMOVE_OL.Name].Value = true;
                    else
                        dgvOutlook.Rows[indexOutlook].Cells[REMOVE_OL.Name].Value = false;
                }
            }
        }

        /// <summary>
        /// Unselect all entries in Outlook Data Grid
        /// </summary>
        private void UnselectAllOutlook()
        {
            for (int index = 0; index < dgvOutlook.Rows.Count; index++)
            {
                dgvOutlook.Rows[index].Cells[REMOVE_OL.Name].Value = false;
            }
        }

        /// Loads data from Outlook (TimeZones, Categories, Current Schedule
		/// </summary>
		protected void LoadOutlookControls()
		{
			//  Don't load the Outlook Schedule yet.
			//  LoadOutlookControls will do that
			outlookControlsInitializing = true;
			dtpOutlook.Value = DateTime.Now;
			OutlookCalendar = dtpOutlook.Value;
			outlookControlsInitializing = false;

			LoadTimeZones();
			LoadOutlookCategories();
			LoadOutlookSchedule();
		}

		/// <summary>
		/// Loads the time zones from the Windows Registry
		/// </summary>
		protected void LoadTimeZones()
		{
			cmbxTimeZones.Items.Clear();
			TimeZoneInfo[] tziWindows = TimeZoneInfo.GetTimeZonesFromRegistry();

			int indexTimeZones;
			cmbxTimeZones.BeginUpdate();
			for (indexTimeZones = tziWindows.GetLowerBound(0); indexTimeZones <= tziWindows.GetUpperBound(0);
				indexTimeZones++)
			{
				cmbxTimeZones.Items.Add(tziWindows[indexTimeZones].DisplayName);
			}
			cmbxTimeZones.EndUpdate();

			indexTimeZones = cmbxTimeZones.FindString(Properties.Resources.ID_TZ_DEFAULT);
			if (indexTimeZones > -1)
				cmbxTimeZones.SelectedIndex = indexTimeZones;
			//cmbxTimeZones.SelectedValue = Properties.Resources.ID_TZ_DEFAULT;
		}

		/// <summary>
		/// Fills the Checkbox Listbox with Categories from Outloook
		/// </summary>
		protected void LoadOutlookCategories()
		{
			chklbOutlookCategories.Items.Clear();
			InteropOutlook.ApplicationClass applOutlook = null;
			try
			{
				applOutlook = new Microsoft.Office.Interop.Outlook.ApplicationClass();
				InteropOutlook.NameSpaceClass nmOutlook = (InteropOutlook.NameSpaceClass)applOutlook.GetNamespace(Properties.Resources.OUTLOOK_MAPI);
				if (nmOutlook.Categories.Count > 0)
				{
					int index;

					bool selected;
					chklbOutlookCategories.BeginUpdate();
					for (index = 1; index <= nmOutlook.Categories.Count; index++)
					{
						InteropOutlook.CategoryClass catAppointments =
							(InteropOutlook.CategoryClass)nmOutlook.Categories[index];
						selected = (catAppointments.Name == Properties.Resources.ID_DEFALT_CATEGORY);
						chklbOutlookCategories.Items.Add(catAppointments.Name,
						(selected ? CheckState.Checked : CheckState.Unchecked));
					}
					chklbOutlookCategories.EndUpdate();
					chklbOutlookCategories.SelectedIndex =
						chklbOutlookCategories.FindString(Properties.Resources.ID_DEFALT_CATEGORY);
				}
			}
			catch (COMException comExp)
			{
				SuspendTimer();
				MessageBox.Show(comExp.Message + comExp.StackTrace, Properties.Resources.ERR_COM_EXCEPTION,
					MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1, (MessageBoxOptions)0);
				ResumeTimer();
				if (Properties.Settings.Default.CopyExceptionsToClipboard)
					Clipboard.SetText(comExp.Message + CRLF + comExp.StackTrace, TextDataFormat.Text);
			}
			finally
			{
				applOutlook = null;
			}
		}

		/// <summary>
		/// Loads the Calendar entries from Outlook based on the selected date + LookAheadWeeks (from the Settings)
		/// weeks and categories selected
		/// </summary>
		protected void LoadOutlookSchedule()
		{
			dgvOutlook.Rows.Clear();

			InteropOutlook.ApplicationClass outlook = null;
			InteropOutlook.NameSpace nmOutlook = null;
			InteropOutlook.Folder olCalendarFolder = null;

			try
			{
				outlook = new Microsoft.Office.Interop.Outlook.ApplicationClass();

				DateTime dtStart = dtpOutlook.Value;
				dtStart = dtStart.Date;
				const int daysInWeek = 7;
				//	Set an end date x weeks from the Application Specified Setting of LookAheadWeeks
				DateTime dtEnd = dtStart.AddDays(daysInWeek * Properties.Settings.Default.LookAheadWeeks);
				string filterDateSearchRange = String.Format(Properties.Resources.FMT_DATE_FILTER,
					dtStart.ToString("g", CultureInfo.CurrentCulture),
					dtEnd.ToString("g", CultureInfo.CurrentCulture));

				StringBuilder filterCategories = new StringBuilder();

				string categories = GetSelectedCategories();
				//	Multiple categories will be checked and separated by an OR
				if (categories.Length > 0)
				{
					string[] category = categories.Split(';');
					int indexCategories;
					int maxCategories = category.GetUpperBound(0);
					int lowCategories = category.GetLowerBound(0);

					for (indexCategories = lowCategories; indexCategories <= maxCategories; indexCategories++)
					{
						filterCategories.Append(Properties.Resources.FILTER_CATEGORIES + category[indexCategories]);
						//  If not the only category and not the last category
						if ((lowCategories != maxCategories) && (indexCategories < maxCategories))
						{
							filterCategories.Append(Properties.Resources.OR);
						}
					}
				}

				string filterCalendar = filterDateSearchRange;
				//	Put the date range search and categories search together
				if (filterCategories.Length > 0)
				{
					filterCalendar += Properties.Resources.AND + "(" + filterCategories.ToString() + ")";
				}

				filterCategories = null;

				nmOutlook = outlook.GetNamespace(Properties.Resources.OUTLOOK_MAPI);
				//  Ralph Hightower - 20071104
				//  FolderClass, ItemClass, and AppointmentItemClass do not appear to work
				//  Use Folder, Item, and AppointmentItem instead
				//InteropOutlook.FolderClass olCalendarFolder = nmOutlook.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderCalendar)
				//    as InteropOutlook.FolderClass;
				olCalendarFolder = nmOutlook.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderCalendar)
				as InteropOutlook.Folder;

				dgvOutlook.SuspendLayout();
				if (olCalendarFolder != null)
				{
					//InteropOutlook.ItemsClass calendarItems = (InteropOutlook.ItemsClass)olCalendarFolder.Items.Restrict(filterCalendar);
					InteropOutlook.Items calendarItems = (InteropOutlook.ItemsClass)olCalendarFolder.Items.Restrict(filterCalendar);
					calendarItems.Sort(Properties.Resources.OUTLOOK_START, Type.Missing);
					foreach (InteropOutlook.AppointmentItem apptItem in calendarItems)
					{
						dgvOutlook.Rows.Add(false, apptItem.Start, apptItem.End, apptItem.Subject, apptItem.Location);
					}
				}
			}
			catch (COMException comExp)
			{
				SuspendTimer();
				MessageBox.Show(comExp.Message + comExp.StackTrace, Properties.Resources.ERR_COM_EXCEPTION,
					MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1, (MessageBoxOptions)0);
				ResumeTimer();
				if (Properties.Settings.Default.CopyExceptionsToClipboard)
					Clipboard.SetText(comExp.Message + CRLF + comExp.StackTrace, TextDataFormat.Text);
			}
			finally
			{
				dgvOutlook.ResumeLayout();
				dgvOutlook.Refresh();
				olCalendarFolder = null;
				nmOutlook = null;
				outlook = null;
			}
		}

        /// <summary>
        /// Updates the Outlook Schedule with the new Excel Schedule
        /// </summary>
        private void UpdateNewSchedule()
        {
			bool somethingToDo = false;

			if (OpenNasaTvSchedule())
			{
                Busy(Properties.Resources.ID_BUSY_UPDATING_SCHEDULE);
                StartTimer();
                somethingToDo = (dgvExcelSchedule.Rows.Count > 0);
				if (somethingToDo)
				{
					dtpOutlook.Value = (DateTime)dgvExcelSchedule.Rows[0].Cells[BEGIN_DATE_TV.Name].Value;
					dtpOutlook.Refresh();
					LoadOutlookSchedule();
				}

				if (somethingToDo)
				{
					SmartSelect();
					RemoveOutlookEntries();
					SelectAllExcel();
					TransferExcelToOutlook();
				}
                Ready();
                StopTimer(Properties.Resources.TIMER_ELAPSED_TIME_NEWSCHEDULEUPDATE);
            }
        }

		/// <summary>
		/// Gets selected Outlook Categories separated by a semicolon
		/// </summary>
		/// <returns>Semi-colon delimited string of Outlook Categories selected in CheckedListBox</returns>
		private string GetSelectedCategories()
		{
			StringBuilder selectedCategories = new StringBuilder();

			CheckedListBox.CheckedItemCollection selectedItems = chklbOutlookCategories.CheckedItems;
			int indexCategories;
			for (indexCategories = 0; indexCategories < selectedItems.Count; indexCategories++)
			{
				if ((indexCategories > 0) && (indexCategories < selectedItems.Count))
					selectedCategories.Append(";");
				selectedCategories.Append(selectedItems[indexCategories].ToString());
			}

			string categories = selectedCategories.ToString();
			selectedCategories = null;
			return (categories);
		}

		/// <summary>
		/// Deletes Calendar entries in Outlook based on schedule entries in Outlook Data Grid that are selected
		/// </summary>
		private void RemoveOutlookEntries()
		{
			InteropOutlook.ApplicationClass outlook = new Microsoft.Office.Interop.Outlook.ApplicationClass();

			int indexGrid;
			SetupProgressBar(dgvOutlook.RowCount);
			for (indexGrid = 0; indexGrid < dgvOutlook.Rows.Count; indexGrid++)
			{
				bool remove = (bool)dgvOutlook.Rows[indexGrid].Cells[REMOVE_OL.Name].Value;
				if (remove)
					RemoveAppointment((DateTime)dgvOutlook.Rows[indexGrid].Cells[BEGIN_DATE_OL.Name].Value,
						(DateTime)dgvOutlook.Rows[indexGrid].Cells[END_DATE_OL.Name].Value,
						(string)dgvOutlook.Rows[indexGrid].Cells[SUBJECT_OL.Name].Value,
						(string)dgvOutlook.Rows[indexGrid].Cells[SITE_OL.Name].Value,
						outlook);
				ProgressBar = indexGrid;
			}
			ProgressBar = -1;

			outlook = null;
		}

		/// <summary>
		/// Adds selected Calendar entries from the Nasa Sts TV Schedule Data Grid that are selected
		/// </summary>
		private void TransferExcelToOutlook()
		{
			InteropOutlook.ApplicationClass outlook = new Microsoft.Office.Interop.Outlook.ApplicationClass();

			string selectedCategories = GetSelectedCategories();

			int gridIndex;
			int maxEntries = dgvExcelSchedule.Rows.Count;
			int itemsAdded = 0;

			DataGridViewRow dgvrScheduleEntry;
			try
			{
				bool transfer = false;
				SetupProgressBar(maxEntries);
				for (gridIndex = 0; gridIndex < maxEntries; gridIndex++)
				{
					transfer = (bool)dgvExcelSchedule.Rows[gridIndex].Cells[ID_ADD.Name].Value;
					if (transfer)
					{
						dgvrScheduleEntry = dgvExcelSchedule.Rows[gridIndex];
						NasaStsTVScheduleEntry nasaStsTVScheduleEntry = new NasaStsTVScheduleEntry(
							(DateTime)dgvrScheduleEntry.Cells[BEGIN_DATE_TV.Name].Value,
							(DateTime)dgvrScheduleEntry.Cells[END_DATE_TV.Name].Value,
							(bool)dgvrScheduleEntry.Cells[CHANGED_TV.Name].Value,
							(string)dgvrScheduleEntry.Cells[SUBJECT_TV.Name].Value,
							0,
							(string)dgvrScheduleEntry.Cells[SITE_TV.Name].Value,
							"",
							ScheduleType.scheduleEntry);
						bool reminder = (bool)dgvrScheduleEntry.Cells[REMINDER_TV.Name].Value;

						AddAppointment(nasaStsTVScheduleEntry, reminder, selectedCategories, outlook);
						itemsAdded++;
					}
					ProgressBar = gridIndex;
				}
				ProgressBar = -1;
			}

			catch (COMException comExp)
			{
				SuspendTimer();
				MessageBox.Show(comExp.Message + comExp.StackTrace, Properties.Resources.ERR_COM_EXCEPTION,
					MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1, (MessageBoxOptions)0);
				ResumeTimer();
				if (Properties.Settings.Default.CopyExceptionsToClipboard)
                    Clipboard.SetText(comExp.Message + CRLF + comExp.StackTrace, TextDataFormat.Text);
			}

			finally
			{
				outlook = null;
				if (itemsAdded > 0)
					LoadOutlookSchedule();
			}
		}

		/// <summary>
		/// Adds the Appointment to the Outlook Calendar
		/// </summary>
		/// <param name="nasaTVSchedule">Class containing the information for the appointment item</param>
		/// <param name="reminder">Set a reminder if true</param>
		/// <param name="categories">Outlook categories to file this appointment under</param>
		/// <param name="outlook">The Outlook application</param>
		private void AddAppointment(NasaStsTVScheduleEntry nasaTVSchedule, bool reminder, string categories,
			InteropOutlook.ApplicationClass outlook)
		{
			try
			{
				string selectedCategories = categories.Replace(";", ", ");

				InteropOutlook.AppointmentItem appt =
					outlook.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem)
					as InteropOutlook.AppointmentItem;
				appt.Start = nasaTVSchedule.BeginDate;
				appt.End = nasaTVSchedule.EndDate;
				appt.Subject = nasaTVSchedule.Subject;
				appt.Location = nasaTVSchedule.Site;
				appt.BusyStatus = Microsoft.Office.Interop.Outlook.OlBusyStatus.olFree;
				appt.Categories = selectedCategories;
				appt.ReminderSet = reminder;
				if (reminder)
					appt.ReminderMinutesBeforeStart = 15;
				appt.Importance = Microsoft.Office.Interop.Outlook.OlImportance.olImportanceNormal;
				appt.BusyStatus = Microsoft.Office.Interop.Outlook.OlBusyStatus.olFree;

				appt.Save();
			}
			catch (COMException comExp)
			{
				SuspendTimer();
				MessageBox.Show(comExp.Message + comExp.StackTrace, Properties.Resources.ERR_COM_EXCEPTION,
					MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1, (MessageBoxOptions)0);
				ResumeTimer();
				if (Properties.Settings.Default.CopyExceptionsToClipboard)
                    Clipboard.SetText(comExp.Message + CRLF + comExp.StackTrace, TextDataFormat.Text);
			}
		}

		/// <summary>
		/// Deletes the Appointment from the Calendar
		/// </summary>
		/// <param name="dtStart">Start Time of the Appointment</param>
		/// <param name="dtEnd">End Time of the Appointment</param>
		/// <param name="subject">Subject of the Appointment</param>
		/// <param name="site">Site of the Appointment</param>
		/// <param name="outlook">Outlook Application to avoid opening and closing repeatedly</param>
		private void RemoveAppointment(DateTime dtStart, DateTime dtEnd, string subject, string site,
			InteropOutlook.ApplicationClass outlook)
		{
			//
			//	COM Exception cause: Single quotes in Subject causes RemoveAppointment to get a COM Exception in Calendar.Items.Restrict(filterAppt)
			//
			string filterAppt = String.Format(Properties.Resources.FMT_FILTER_DATE_SUBJECT_LOCATION,
				dtStart.ToString("g", CultureInfo.CurrentCulture),
				dtEnd.ToString("g", CultureInfo.CurrentCulture),
				subject.Replace("'", "''"),
				site.Replace("'", "''"));

			InteropOutlook.NameSpace nmOutlook = null;
			InteropOutlook.Folder olCalendarFolder = null;
			try
			{
				nmOutlook = outlook.GetNamespace("MAPI");
				//  Ralph Hightower - 20071104
				//  FolderClass, ItemClass, and AppointmentItemClass do not appear to work
				//  Use Folder, Item, and AppointmentItem instead
				//  InteropOutlook.FolderClass olCalendarFolder = nmOutlook.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderCalendar)
				//    as InteropOutlook.FolderClass;
				olCalendarFolder = nmOutlook.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderCalendar)
				as InteropOutlook.Folder;

				if (olCalendarFolder != null)
				{
					//InteropOutlook.ItemsClass calendarItems = (InteropOutlook.ItemsClass)olCalendarFolder.Items.Restrict(filterCalendar);
					InteropOutlook.Items calendarItems = (InteropOutlook.ItemsClass)olCalendarFolder.Items.Restrict(filterAppt);
					calendarItems.Sort(Properties.Resources.OUTLOOK_START, Type.Missing);
					foreach (InteropOutlook.AppointmentItem apptItem in calendarItems)
					{
						apptItem.Delete();
					}
				}
			}
			catch (COMException comExp)
			{
				SuspendTimer();
				MessageBox.Show(comExp.Message + comExp.StackTrace, Properties.Resources.ERR_COM_EXCEPTION,
					MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1, (MessageBoxOptions)0);
				ResumeTimer();
				if (Properties.Settings.Default.CopyExceptionsToClipboard)
                    Clipboard.SetText(comExp.Message + CRLF + comExp.StackTrace, TextDataFormat.Text);
			}
			finally
			{
				olCalendarFolder = null;
				nmOutlook = null;
			}
		}

		/// <summary>
		/// Select a group of files to load in bulk array of filenames must be sorted by filename
		///
		/// The sequence is (after loading the first file and transferring to Outlook):
		///
		/// Read and process a schedule file
		/// Select those entries in Outlook that are in the time span of the schedule file just process
		/// Remove those entries from Outlook
		/// Transfer the entries from the file just processed to Outlook
		/// </summary>
		/// <param name="filenames">Group of NASA TV Schedule files to import</param>
		private void ImportMultipleSchedules(string[] filenames)
		{
			int indexFiles;
			for (indexFiles = filenames.GetLowerBound(0); indexFiles <= filenames.GetUpperBound(0); indexFiles++)
			{
				FileInfo fiExcelSchedule = new FileInfo(filenames[indexFiles]);
				lblNasaStsTVScheduleFile.Text = fiExcelSchedule.Name;
				lblNasaStsTVScheduleFile.Refresh();
				fiExcelSchedule = null;

				string excelFile = filenames[indexFiles];

				Busy(Properties.Resources.ID_BULKIMPORT_READFILE + excelFile);
				LoadExcelSchedule(filenames[indexFiles]);
				if (indexFiles == filenames.GetLowerBound(0))
				{
					if (dgvExcelSchedule.RowCount > 0)
					{
						dtpOutlook.Value = (DateTime)dgvExcelSchedule.Rows[0].Cells[BEGIN_DATE_TV.Name].Value;
						dtpOutlook.Refresh();
                        LoadOutlookSchedule();
					}
					Busy(Properties.Resources.ID_BULKIMPORT_SELECTALLEXCEL);
					SelectAllExcel();
					Busy(Properties.Resources.ID_BULKIMPORT_TRANSFERTOOUTLOOK);
					TransferExcelToOutlook();
				}
				else
				{
					Busy(Properties.Resources.ID_BULKIMPORT_SMARTSELECT);
					SmartSelect();
					Busy(Properties.Resources.ID_BULKIMPORT_REMOVINGOUTLOOKENTRIES);
					RemoveOutlookEntries();
					Busy(Properties.Resources.ID_BULKIMPORT_SELECTALLEXCEL);
					SelectAllExcel();
					Busy(Properties.Resources.ID_BULKIMPORT_TRANSFERTOOUTLOOK);
					TransferExcelToOutlook();
				}
			}
		}

		#endregion

        private void MainForm_AutoSizeChanged(object sender, EventArgs e)
        {

        }

        private void MainForm_MaximizedBoundsChanged(object sender, EventArgs e)
        {

        }

        private void MainForm_MaximumSizeChanged(object sender, EventArgs e)
        {

        }

        private void MainForm_MinimumSizeChanged(object sender, EventArgs e)
        {

        }

        private void MainForm_SizeChanged(object sender, EventArgs e)
        {

        }
        
    }
}