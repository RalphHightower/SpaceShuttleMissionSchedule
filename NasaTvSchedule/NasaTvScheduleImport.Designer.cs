namespace PermanentVacations.Nasa.Sts.OutlookCalendar
{
    partial class MainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
			this.btnOpenNasaTvSchedule = new System.Windows.Forms.Button();
			this.dgvExcelSchedule = new System.Windows.Forms.DataGridView();
			this.ID_ADD = new System.Windows.Forms.DataGridViewCheckBoxColumn();
			this.ORBIT_TV = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.BEGIN_DATE_TV = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.END_DATE_TV = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.REMINDER_TV = new System.Windows.Forms.DataGridViewCheckBoxColumn();
			this.CHANGED_TV = new System.Windows.Forms.DataGridViewCheckBoxColumn();
			this.SUBJECT_TV = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.SITE_TV = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.chklbOutlookCategories = new System.Windows.Forms.CheckedListBox();
			this.dgvOutlook = new System.Windows.Forms.DataGridView();
			this.REMOVE_OL = new System.Windows.Forms.DataGridViewCheckBoxColumn();
			this.BEGIN_DATE_OL = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.END_DATE_OL = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.SUBJECT_OL = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.SITE_OL = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.btnRemoveMarkedEntries = new System.Windows.Forms.Button();
			this.btnTransferTVSchedule = new System.Windows.Forms.Button();
			this.btnExitApplication = new System.Windows.Forms.Button();
			this.ofExcelSchedule = new System.Windows.Forms.OpenFileDialog();
			this.chkInteropExcel = new System.Windows.Forms.CheckBox();
			this.chkDocked = new System.Windows.Forms.CheckBox();
			this.chkOrbit = new System.Windows.Forms.CheckBox();
			this.btnSelectAllExcel = new System.Windows.Forms.Button();
			this.btnSelectAllOutlook = new System.Windows.Forms.Button();
			this.btnUnselectAllExcel = new System.Windows.Forms.Button();
			this.btnUnselectAllOutlook = new System.Windows.Forms.Button();
			this.cmbxTimeZones = new System.Windows.Forms.ComboBox();
			this.btnRefreshOutlookCategories = new System.Windows.Forms.Button();
			this.statusStrip = new System.Windows.Forms.StatusStrip();
			this.toolStripProgressBar = new System.Windows.Forms.ToolStripProgressBar();
			this.toolStripStatusLabel = new System.Windows.Forms.ToolStripStatusLabel();
			this.label1 = new System.Windows.Forms.Label();
			this.lblNasaStsTVScheduleFile = new System.Windows.Forms.Label();
			this.btnSmartSelect = new System.Windows.Forms.Button();
			this.btnBulkImport = new System.Windows.Forms.Button();
			this.menuStrip = new System.Windows.Forms.MenuStrip();
			this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.newToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.openToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.toolStripSeparator = new System.Windows.Forms.ToolStripSeparator();
			this.saveToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.saveAsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
			this.printToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.printPreviewToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
			this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.editToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.undoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.redoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.toolStripSeparator3 = new System.Windows.Forms.ToolStripSeparator();
			this.cutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.copyToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.pasteToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.toolStripSeparator4 = new System.Windows.Forms.ToolStripSeparator();
			this.selectAllToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.toolsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.customizeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.optionsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.contentsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.indexToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.searchToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.toolStripSeparator5 = new System.Windows.Forms.ToolStripSeparator();
			this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.dtpOutlook = new System.Windows.Forms.DateTimePicker();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.dgvExcelSchedule)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dgvOutlook)).BeginInit();
			this.SuspendLayout();
			// 
			// btnOpenNasaTvSchedule
			// 
			this.btnOpenNasaTvSchedule.AccessibleDescription = "Open Nasa Shuttle TV Schedule";
			this.btnOpenNasaTvSchedule.AccessibleName = "Nasa TV Schedule";
			this.btnOpenNasaTvSchedule.Location = new System.Drawing.Point(11, 40);
			this.btnOpenNasaTvSchedule.Name = "btnOpenNasaTvSchedule";
			this.btnOpenNasaTvSchedule.Size = new System.Drawing.Size(147, 23);
			this.btnOpenNasaTvSchedule.TabIndex = 0;
			this.btnOpenNasaTvSchedule.Text = "Open Nasa TV Schedule";
			this.btnOpenNasaTvSchedule.UseVisualStyleBackColor = true;
			this.btnOpenNasaTvSchedule.MouseLeave += new System.EventHandler(this.btnOpenNasaTvSchedule_MouseLeave);
			this.btnOpenNasaTvSchedule.Click += new System.EventHandler(this.btnOpenNasaTvSchedule_Click);
			this.btnOpenNasaTvSchedule.MouseHover += new System.EventHandler(this.btnOpenNasaTvSchedule_MouseHover);
			// 
			// dgvExcelSchedule
			// 
			this.dgvExcelSchedule.AccessibleDescription = "Nasa TV Schedule from Excel";
			this.dgvExcelSchedule.AccessibleName = "Nasa TV Schedule Entries";
			this.dgvExcelSchedule.AllowUserToAddRows = false;
			this.dgvExcelSchedule.CausesValidation = false;
			this.dgvExcelSchedule.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.dgvExcelSchedule.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ID_ADD,
            this.ORBIT_TV,
            this.BEGIN_DATE_TV,
            this.END_DATE_TV,
            this.REMINDER_TV,
            this.CHANGED_TV,
            this.SUBJECT_TV,
            this.SITE_TV});
			this.dgvExcelSchedule.Location = new System.Drawing.Point(11, 94);
			this.dgvExcelSchedule.Name = "dgvExcelSchedule";
			this.dgvExcelSchedule.Size = new System.Drawing.Size(767, 138);
			this.dgvExcelSchedule.TabIndex = 1;
			this.dgvExcelSchedule.MouseHover += new System.EventHandler(this.dgvExcelSchedule_MouseHover);
			this.dgvExcelSchedule.MouseLeave += new System.EventHandler(this.dgvExcelSchedule_MouseLeave);
			this.dgvExcelSchedule.CurrentCellDirtyStateChanged += new System.EventHandler(this.dgvExcelSchedule_CurrentCellDirtyStateChanged);
			this.dgvExcelSchedule.DataError += new System.Windows.Forms.DataGridViewDataErrorEventHandler(this.dgvExcelSchedule_DataError);
			// 
			// ID_ADD
			// 
			this.ID_ADD.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader;
			this.ID_ADD.HeaderText = "Add";
			this.ID_ADD.Name = "ID_ADD";
			this.ID_ADD.ToolTipText = "Select this TV Schedule entry to add it to Outlook as an Appointment";
			this.ID_ADD.Width = 32;
			// 
			// ORBIT_TV
			// 
			this.ORBIT_TV.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
			this.ORBIT_TV.HeaderText = "ORBIT";
			this.ORBIT_TV.Name = "ORBIT_TV";
			this.ORBIT_TV.ReadOnly = true;
			this.ORBIT_TV.ToolTipText = "Current orbit of Space Shuttle";
			this.ORBIT_TV.Width = 65;
			// 
			// BEGIN_DATE_TV
			// 
			this.BEGIN_DATE_TV.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
			dataGridViewCellStyle5.Format = "g";
			dataGridViewCellStyle5.NullValue = null;
			this.BEGIN_DATE_TV.DefaultCellStyle = dataGridViewCellStyle5;
			this.BEGIN_DATE_TV.HeaderText = "BEGIN DATE";
			this.BEGIN_DATE_TV.Name = "BEGIN_DATE_TV";
			this.BEGIN_DATE_TV.ReadOnly = true;
			this.BEGIN_DATE_TV.ToolTipText = "Beginning date and time for Nasa TV Schedule";
			this.BEGIN_DATE_TV.Width = 97;
			// 
			// END_DATE_TV
			// 
			this.END_DATE_TV.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
			this.END_DATE_TV.HeaderText = "END DATE";
			this.END_DATE_TV.Name = "END_DATE_TV";
			this.END_DATE_TV.ReadOnly = true;
			this.END_DATE_TV.ToolTipText = "Ending date and time for Nasa TV entry";
			this.END_DATE_TV.Width = 87;
			// 
			// REMINDER_TV
			// 
			this.REMINDER_TV.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
			this.REMINDER_TV.HeaderText = "Reminder";
			this.REMINDER_TV.Name = "REMINDER_TV";
			this.REMINDER_TV.ToolTipText = "If adding this entry to Outlook, add it as a reminder";
			this.REMINDER_TV.Width = 58;
			// 
			// CHANGED_TV
			// 
			this.CHANGED_TV.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
			this.CHANGED_TV.HeaderText = "CHANGED";
			this.CHANGED_TV.Name = "CHANGED_TV";
			this.CHANGED_TV.ReadOnly = true;
			this.CHANGED_TV.Resizable = System.Windows.Forms.DataGridViewTriState.True;
			this.CHANGED_TV.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
			this.CHANGED_TV.ToolTipText = "Nasa changed this TV schedule entry in this revision";
			this.CHANGED_TV.Width = 85;
			// 
			// SUBJECT_TV
			// 
			this.SUBJECT_TV.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
			this.SUBJECT_TV.HeaderText = "SUBJECT";
			this.SUBJECT_TV.Name = "SUBJECT_TV";
			this.SUBJECT_TV.ReadOnly = true;
			this.SUBJECT_TV.ToolTipText = "Subject for this Nasa TV schedule entry";
			this.SUBJECT_TV.Width = 80;
			// 
			// SITE_TV
			// 
			this.SITE_TV.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
			this.SITE_TV.HeaderText = "SITE";
			this.SITE_TV.Name = "SITE_TV";
			this.SITE_TV.ToolTipText = "Location for this TV Schedule entry";
			this.SITE_TV.Width = 56;
			// 
			// chklbOutlookCategories
			// 
			this.chklbOutlookCategories.AccessibleDescription = "Select the Outlook categories to include in the Outlook appointment item grid";
			this.chklbOutlookCategories.AccessibleName = "Outlook Categories";
			this.chklbOutlookCategories.CheckOnClick = true;
			this.chklbOutlookCategories.FormattingEnabled = true;
			this.chklbOutlookCategories.Location = new System.Drawing.Point(124, 266);
			this.chklbOutlookCategories.Name = "chklbOutlookCategories";
			this.chklbOutlookCategories.Size = new System.Drawing.Size(200, 79);
			this.chklbOutlookCategories.Sorted = true;
			this.chklbOutlookCategories.TabIndex = 3;
			this.chklbOutlookCategories.MouseHover += new System.EventHandler(this.chklbOutlookCategories_MouseHover);
			this.chklbOutlookCategories.MouseLeave += new System.EventHandler(this.chklbOutlookCategories_MouseLeave);
			this.chklbOutlookCategories.SelectedValueChanged += new System.EventHandler(this.chklbOutlookCategories_SelectedValueChanged);
			// 
			// dgvOutlook
			// 
			this.dgvOutlook.AccessibleDescription = "Appointment items from Outlook";
			this.dgvOutlook.AccessibleName = "Outlook Appointment Entries";
			this.dgvOutlook.AllowUserToAddRows = false;
			this.dgvOutlook.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.dgvOutlook.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.REMOVE_OL,
            this.BEGIN_DATE_OL,
            this.END_DATE_OL,
            this.SUBJECT_OL,
            this.SITE_OL});
			this.dgvOutlook.Location = new System.Drawing.Point(11, 351);
			this.dgvOutlook.Name = "dgvOutlook";
			this.dgvOutlook.Size = new System.Drawing.Size(767, 138);
			this.dgvOutlook.TabIndex = 5;
			this.dgvOutlook.MouseHover += new System.EventHandler(this.dgvOutlook_MouseHover);
			this.dgvOutlook.MouseLeave += new System.EventHandler(this.dgvOutlook_MouseLeave);
			this.dgvOutlook.CurrentCellDirtyStateChanged += new System.EventHandler(this.dgvOutlook_CurrentCellDirtyStateChanged);
			this.dgvOutlook.DataError += new System.Windows.Forms.DataGridViewDataErrorEventHandler(this.dgvOutlook_DataError);
			// 
			// REMOVE_OL
			// 
			this.REMOVE_OL.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader;
			this.REMOVE_OL.HeaderText = "Remove";
			this.REMOVE_OL.Name = "REMOVE_OL";
			this.REMOVE_OL.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
			this.REMOVE_OL.ToolTipText = "Select this Appointment to delete it from Outlook";
			this.REMOVE_OL.Width = 72;
			// 
			// BEGIN_DATE_OL
			// 
			this.BEGIN_DATE_OL.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
			dataGridViewCellStyle6.Format = "g";
			dataGridViewCellStyle6.NullValue = null;
			this.BEGIN_DATE_OL.DefaultCellStyle = dataGridViewCellStyle6;
			this.BEGIN_DATE_OL.HeaderText = "BEGIN DATE";
			this.BEGIN_DATE_OL.Name = "BEGIN_DATE_OL";
			this.BEGIN_DATE_OL.ReadOnly = true;
			this.BEGIN_DATE_OL.ToolTipText = "Beginning Date & Time for this Appointment";
			this.BEGIN_DATE_OL.Width = 97;
			// 
			// END_DATE_OL
			// 
			this.END_DATE_OL.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
			this.END_DATE_OL.HeaderText = "END DATE";
			this.END_DATE_OL.Name = "END_DATE_OL";
			this.END_DATE_OL.ReadOnly = true;
			this.END_DATE_OL.ToolTipText = "Ending date and time for this Appointment item";
			this.END_DATE_OL.Width = 87;
			// 
			// SUBJECT_OL
			// 
			this.SUBJECT_OL.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
			this.SUBJECT_OL.HeaderText = "SUBJECT";
			this.SUBJECT_OL.Name = "SUBJECT_OL";
			this.SUBJECT_OL.ReadOnly = true;
			this.SUBJECT_OL.ToolTipText = "Subject of this Appointment item";
			this.SUBJECT_OL.Width = 80;
			// 
			// SITE_OL
			// 
			this.SITE_OL.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
			this.SITE_OL.HeaderText = "SITE";
			this.SITE_OL.Name = "SITE_OL";
			this.SITE_OL.ReadOnly = true;
			this.SITE_OL.ToolTipText = "Location for this Appointment item";
			this.SITE_OL.Width = 56;
			// 
			// btnRemoveMarkedEntries
			// 
			this.btnRemoveMarkedEntries.AccessibleDescription = "Remove the selected entries in the Outlook schedule from Outlook";
			this.btnRemoveMarkedEntries.AccessibleName = "Remove Selected Outlook Entries";
			this.btnRemoveMarkedEntries.Location = new System.Drawing.Point(490, 294);
			this.btnRemoveMarkedEntries.Name = "btnRemoveMarkedEntries";
			this.btnRemoveMarkedEntries.Size = new System.Drawing.Size(142, 23);
			this.btnRemoveMarkedEntries.TabIndex = 6;
			this.btnRemoveMarkedEntries.Text = "Remove Selected Entries";
			this.btnRemoveMarkedEntries.UseVisualStyleBackColor = true;
			this.btnRemoveMarkedEntries.MouseLeave += new System.EventHandler(this.btnRemoveMarkedEntries_MouseLeave);
			this.btnRemoveMarkedEntries.Click += new System.EventHandler(this.btnRemoveMarkedEntries_Click);
			this.btnRemoveMarkedEntries.MouseHover += new System.EventHandler(this.btnRemoveMarkedEntries_MouseHover);
			// 
			// btnTransferTVSchedule
			// 
			this.btnTransferTVSchedule.AccessibleDescription = "Transfer the selected Nasa schedule items into Outlook as Appointment items";
			this.btnTransferTVSchedule.AccessibleName = "Transfer TV Schedule";
			this.btnTransferTVSchedule.Location = new System.Drawing.Point(490, 267);
			this.btnTransferTVSchedule.Name = "btnTransferTVSchedule";
			this.btnTransferTVSchedule.Size = new System.Drawing.Size(142, 23);
			this.btnTransferTVSchedule.TabIndex = 7;
			this.btnTransferTVSchedule.Text = "Transfer TV Schedule";
			this.btnTransferTVSchedule.UseVisualStyleBackColor = true;
			this.btnTransferTVSchedule.MouseLeave += new System.EventHandler(this.btnTransferTVSchedule_MouseLeave);
			this.btnTransferTVSchedule.Click += new System.EventHandler(this.btnTransferTVSchedule_Click);
			this.btnTransferTVSchedule.MouseHover += new System.EventHandler(this.btnTransferTVSchedule_MouseHover);
			// 
			// btnExitApplication
			// 
			this.btnExitApplication.AccessibleDescription = "Exit Application";
			this.btnExitApplication.AccessibleName = "Exit Application";
			this.btnExitApplication.Location = new System.Drawing.Point(11, 505);
			this.btnExitApplication.Name = "btnExitApplication";
			this.btnExitApplication.Size = new System.Drawing.Size(142, 23);
			this.btnExitApplication.TabIndex = 8;
			this.btnExitApplication.Text = "Exit Application";
			this.btnExitApplication.UseVisualStyleBackColor = true;
			this.btnExitApplication.MouseLeave += new System.EventHandler(this.btnExitApplication_MouseLeave);
			this.btnExitApplication.Click += new System.EventHandler(this.btnExitApplication_Click);
			this.btnExitApplication.MouseHover += new System.EventHandler(this.btnExitApplication_MouseHover);
			// 
			// ofExcelSchedule
			// 
			this.ofExcelSchedule.DefaultExt = "xls";
			this.ofExcelSchedule.FileName = "*tvsked_rev*.xls";
			this.ofExcelSchedule.Filter = "Excel Spreadsheets|*.xls";
			this.ofExcelSchedule.InitialDirectory = "E:\\Documents and Settings\\RalphHightower\\My Documents\\Nasa\\";
			// 
			// chkInteropExcel
			// 
			this.chkInteropExcel.AccessibleDescription = "Developer use only: Do not click this.";
			this.chkInteropExcel.AccessibleName = "Developer Use Only";
			this.chkInteropExcel.AutoSize = true;
			this.chkInteropExcel.Location = new System.Drawing.Point(621, 71);
			this.chkInteropExcel.Name = "chkInteropExcel";
			this.chkInteropExcel.Size = new System.Drawing.Size(106, 17);
			this.chkInteropExcel.TabIndex = 9;
			this.chkInteropExcel.Text = "Try Excel Interop";
			this.chkInteropExcel.UseVisualStyleBackColor = true;
			this.chkInteropExcel.Visible = false;
			// 
			// chkDocked
			// 
			this.chkDocked.AccessibleDescription = "Space Shuttle is docked at ISS (will not be reliable in revised Nasa schedules)";
			this.chkDocked.AccessibleName = "Docked Status";
			this.chkDocked.AutoSize = true;
			this.chkDocked.Location = new System.Drawing.Point(164, 69);
			this.chkDocked.Name = "chkDocked";
			this.chkDocked.Size = new System.Drawing.Size(64, 17);
			this.chkDocked.TabIndex = 10;
			this.chkDocked.Text = "Docked";
			this.chkDocked.UseVisualStyleBackColor = true;
			this.chkDocked.MouseLeave += new System.EventHandler(this.chkDocked_MouseLeave);
			this.chkDocked.MouseHover += new System.EventHandler(this.chkDocked_MouseHover);
			// 
			// chkOrbit
			// 
			this.chkOrbit.AccessibleDescription = "Shuttle is in orbit";
			this.chkOrbit.AccessibleName = "Orbit Status";
			this.chkOrbit.AutoSize = true;
			this.chkOrbit.Location = new System.Drawing.Point(241, 69);
			this.chkOrbit.Name = "chkOrbit";
			this.chkOrbit.Size = new System.Drawing.Size(48, 17);
			this.chkOrbit.TabIndex = 11;
			this.chkOrbit.Text = "Orbit";
			this.chkOrbit.UseVisualStyleBackColor = true;
			this.chkOrbit.MouseLeave += new System.EventHandler(this.chkOrbit_MouseLeave);
			this.chkOrbit.MouseHover += new System.EventHandler(this.chkOrbit_MouseHover);
			// 
			// btnSelectAllExcel
			// 
			this.btnSelectAllExcel.AccessibleDescription = "Select all entries in the Nasa TV Schedule grid to load into Outlook";
			this.btnSelectAllExcel.AccessibleName = "Select All Excel";
			this.btnSelectAllExcel.Location = new System.Drawing.Point(172, 40);
			this.btnSelectAllExcel.Name = "btnSelectAllExcel";
			this.btnSelectAllExcel.Size = new System.Drawing.Size(147, 23);
			this.btnSelectAllExcel.TabIndex = 13;
			this.btnSelectAllExcel.Text = "Select All";
			this.btnSelectAllExcel.UseVisualStyleBackColor = true;
			this.btnSelectAllExcel.MouseLeave += new System.EventHandler(this.btnSelectAllExcel_MouseLeave);
			this.btnSelectAllExcel.Click += new System.EventHandler(this.btnSelectAllExcel_Click);
			this.btnSelectAllExcel.MouseHover += new System.EventHandler(this.btnSelectAllExcel_MouseHover);
			// 
			// btnSelectAllOutlook
			// 
			this.btnSelectAllOutlook.AccessibleDescription = "Select all entries in the Outlook schedule to delete";
			this.btnSelectAllOutlook.AccessibleName = "Select Outlook Entries";
			this.btnSelectAllOutlook.Location = new System.Drawing.Point(338, 267);
			this.btnSelectAllOutlook.Name = "btnSelectAllOutlook";
			this.btnSelectAllOutlook.Size = new System.Drawing.Size(142, 23);
			this.btnSelectAllOutlook.TabIndex = 14;
			this.btnSelectAllOutlook.Text = "Select All";
			this.btnSelectAllOutlook.UseVisualStyleBackColor = true;
			this.btnSelectAllOutlook.MouseLeave += new System.EventHandler(this.btnSelectAllOutlook_MouseLeave);
			this.btnSelectAllOutlook.Click += new System.EventHandler(this.btnSelectAllOutlook_Click);
			this.btnSelectAllOutlook.MouseHover += new System.EventHandler(this.btnSelectAllOutlook_MouseHover);
			// 
			// btnUnselectAllExcel
			// 
			this.btnUnselectAllExcel.AccessibleDescription = "Unselect all Nasa TV schedule entries";
			this.btnUnselectAllExcel.AccessibleName = "Unselect All Excel";
			this.btnUnselectAllExcel.Location = new System.Drawing.Point(333, 40);
			this.btnUnselectAllExcel.Name = "btnUnselectAllExcel";
			this.btnUnselectAllExcel.Size = new System.Drawing.Size(147, 23);
			this.btnUnselectAllExcel.TabIndex = 15;
			this.btnUnselectAllExcel.Text = "Unselect All";
			this.btnUnselectAllExcel.UseVisualStyleBackColor = true;
			this.btnUnselectAllExcel.MouseLeave += new System.EventHandler(this.btnUnselectAllExcel_MouseLeave);
			this.btnUnselectAllExcel.Click += new System.EventHandler(this.btnUnselectAllExcel_Click);
			this.btnUnselectAllExcel.MouseHover += new System.EventHandler(this.btnUnselectAllExcel_MouseHover);
			// 
			// btnUnselectAllOutlook
			// 
			this.btnUnselectAllOutlook.AccessibleDescription = "Unselect all Outlook schedule entries";
			this.btnUnselectAllOutlook.AccessibleName = "Unselect All Outlook Entries";
			this.btnUnselectAllOutlook.Location = new System.Drawing.Point(338, 294);
			this.btnUnselectAllOutlook.Name = "btnUnselectAllOutlook";
			this.btnUnselectAllOutlook.Size = new System.Drawing.Size(142, 23);
			this.btnUnselectAllOutlook.TabIndex = 16;
			this.btnUnselectAllOutlook.Text = "Unselect All";
			this.btnUnselectAllOutlook.UseVisualStyleBackColor = true;
			this.btnUnselectAllOutlook.MouseLeave += new System.EventHandler(this.btnUnselectAllOutlook_MouseLeave);
			this.btnUnselectAllOutlook.Click += new System.EventHandler(this.btnUnselectAllOutlook_Click);
			this.btnUnselectAllOutlook.MouseHover += new System.EventHandler(this.btnUnselectAllOutlook_MouseHover);
			// 
			// cmbxTimeZones
			// 
			this.cmbxTimeZones.AccessibleDescription = "Changes the time zone for the appointment items";
			this.cmbxTimeZones.AccessibleName = "Time Zone Selection";
			this.cmbxTimeZones.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cmbxTimeZones.FormattingEnabled = true;
			this.cmbxTimeZones.Location = new System.Drawing.Point(434, 240);
			this.cmbxTimeZones.Name = "cmbxTimeZones";
			this.cmbxTimeZones.Size = new System.Drawing.Size(344, 21);
			this.cmbxTimeZones.TabIndex = 17;
			this.cmbxTimeZones.MouseHover += new System.EventHandler(this.cmbxTimeZones_MouseHover);
			this.cmbxTimeZones.SelectedIndexChanged += new System.EventHandler(this.cmbxTimeZones_SelectedIndexChanged);
			this.cmbxTimeZones.MouseLeave += new System.EventHandler(this.cmbxTimeZones_MouseLeave);
			// 
			// btnRefreshOutlookCategories
			// 
			this.btnRefreshOutlookCategories.AccessibleDescription = "Reload the Outlook Categories";
			this.btnRefreshOutlookCategories.AccessibleName = "Refresh Outlook Categories";
			this.btnRefreshOutlookCategories.Location = new System.Drawing.Point(490, 321);
			this.btnRefreshOutlookCategories.Name = "btnRefreshOutlookCategories";
			this.btnRefreshOutlookCategories.Size = new System.Drawing.Size(142, 23);
			this.btnRefreshOutlookCategories.TabIndex = 18;
			this.btnRefreshOutlookCategories.Text = "Refresh Categories";
			this.btnRefreshOutlookCategories.UseVisualStyleBackColor = true;
			this.btnRefreshOutlookCategories.MouseLeave += new System.EventHandler(this.btnRefreshOutlookCategories_MouseLeave);
			this.btnRefreshOutlookCategories.Click += new System.EventHandler(this.btnRefreshOutlookCategories_Click);
			this.btnRefreshOutlookCategories.MouseHover += new System.EventHandler(this.btnRefreshOutlookCategories_MouseHover);
			// 
			// statusStrip
			// 
			this.statusStrip.Location = new System.Drawing.Point(0, 551);
			this.statusStrip.Name = "statusStrip";
			this.statusStrip.Size = new System.Drawing.Size(795, 22);
			this.statusStrip.TabIndex = 19;
			// 
			// toolStripProgressBar
			// 
			this.toolStripProgressBar.AccessibleDescription = "Indicates the progress of long operations";
			this.toolStripProgressBar.AccessibleName = "Progress Indicator";
			this.toolStripProgressBar.Name = "toolStripProgressBar";
			this.toolStripProgressBar.Size = new System.Drawing.Size(250, 16);
			// 
			// toolStripStatusLabel
			// 
			this.toolStripStatusLabel.AccessibleDescription = "Indicates if program is busy";
			this.toolStripStatusLabel.AccessibleName = "Busy Indicator";
			this.toolStripStatusLabel.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
			this.toolStripStatusLabel.Margin = new System.Windows.Forms.Padding(10, 3, 0, 2);
			this.toolStripStatusLabel.Name = "toolStripStatusLabel";
			this.toolStripStatusLabel.Size = new System.Drawing.Size(0, 17);
			this.toolStripStatusLabel.ToolTipText = "Busy Indicator";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(11, 24);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(165, 13);
			this.label1.TabIndex = 20;
			this.label1.Text = "STS TV Schedule File Processed";
			// 
			// lblNasaStsTVScheduleFile
			// 
			this.lblNasaStsTVScheduleFile.AutoSize = true;
			this.lblNasaStsTVScheduleFile.Location = new System.Drawing.Point(179, 24);
			this.lblNasaStsTVScheduleFile.Name = "lblNasaStsTVScheduleFile";
			this.lblNasaStsTVScheduleFile.Size = new System.Drawing.Size(136, 13);
			this.lblNasaStsTVScheduleFile.TabIndex = 21;
			this.lblNasaStsTVScheduleFile.Text = "Excel File Name Processed";
			// 
			// btnSmartSelect
			// 
			this.btnSmartSelect.Location = new System.Drawing.Point(338, 321);
			this.btnSmartSelect.Name = "btnSmartSelect";
			this.btnSmartSelect.Size = new System.Drawing.Size(142, 23);
			this.btnSmartSelect.TabIndex = 22;
			this.btnSmartSelect.Text = "Smart Select";
			this.btnSmartSelect.UseVisualStyleBackColor = true;
			this.btnSmartSelect.MouseLeave += new System.EventHandler(this.btnSmartSelect_MouseLeave);
			this.btnSmartSelect.Click += new System.EventHandler(this.btnSmartSelect_Click);
			this.btnSmartSelect.MouseHover += new System.EventHandler(this.btnSmartSelect_MouseHover);
			// 
			// btnBulkImport
			// 
			this.btnBulkImport.Location = new System.Drawing.Point(642, 267);
			this.btnBulkImport.Name = "btnBulkImport";
			this.btnBulkImport.Size = new System.Drawing.Size(147, 23);
			this.btnBulkImport.TabIndex = 23;
			this.btnBulkImport.Text = "Bulk Import";
			this.btnBulkImport.UseVisualStyleBackColor = true;
			this.btnBulkImport.MouseLeave += new System.EventHandler(this.btnBulkImport_MouseLeave);
			this.btnBulkImport.Click += new System.EventHandler(this.btnBulkImport_Click);
			this.btnBulkImport.MouseHover += new System.EventHandler(this.btnBulkImport_MouseHover);
			// 
			// menuStrip
			// 
			this.menuStrip.Location = new System.Drawing.Point(0, 0);
			this.menuStrip.Name = "menuStrip";
			this.menuStrip.Size = new System.Drawing.Size(795, 24);
			this.menuStrip.TabIndex = 24;
			this.menuStrip.Text = "menuStrip";
			// 
			// fileToolStripMenuItem
			// 
			this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.newToolStripMenuItem,
            this.openToolStripMenuItem,
            this.toolStripSeparator,
            this.saveToolStripMenuItem,
            this.saveAsToolStripMenuItem,
            this.toolStripSeparator1,
            this.printToolStripMenuItem,
            this.printPreviewToolStripMenuItem,
            this.toolStripSeparator2,
            this.exitToolStripMenuItem});
			this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
			this.fileToolStripMenuItem.Size = new System.Drawing.Size(35, 20);
			this.fileToolStripMenuItem.Text = "&File";
			// 
			// newToolStripMenuItem
			// 
			this.newToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("newToolStripMenuItem.Image")));
			this.newToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.newToolStripMenuItem.Name = "newToolStripMenuItem";
			this.newToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.N)));
			this.newToolStripMenuItem.Size = new System.Drawing.Size(139, 22);
			this.newToolStripMenuItem.Text = "&New";
			// 
			// openToolStripMenuItem
			// 
			this.openToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("openToolStripMenuItem.Image")));
			this.openToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.openToolStripMenuItem.Name = "openToolStripMenuItem";
			this.openToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.O)));
			this.openToolStripMenuItem.Size = new System.Drawing.Size(139, 22);
			this.openToolStripMenuItem.Text = "&Open";
			this.openToolStripMenuItem.Click += new System.EventHandler(this.openToolStripMenuItem_Click);
			// 
			// toolStripSeparator
			// 
			this.toolStripSeparator.Name = "toolStripSeparator";
			this.toolStripSeparator.Size = new System.Drawing.Size(136, 6);
			// 
			// saveToolStripMenuItem
			// 
			this.saveToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("saveToolStripMenuItem.Image")));
			this.saveToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.saveToolStripMenuItem.Name = "saveToolStripMenuItem";
			this.saveToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.S)));
			this.saveToolStripMenuItem.Size = new System.Drawing.Size(139, 22);
			this.saveToolStripMenuItem.Text = "&Save";
			// 
			// saveAsToolStripMenuItem
			// 
			this.saveAsToolStripMenuItem.Name = "saveAsToolStripMenuItem";
			this.saveAsToolStripMenuItem.Size = new System.Drawing.Size(139, 22);
			this.saveAsToolStripMenuItem.Text = "Save &As";
			// 
			// toolStripSeparator1
			// 
			this.toolStripSeparator1.Name = "toolStripSeparator1";
			this.toolStripSeparator1.Size = new System.Drawing.Size(136, 6);
			// 
			// printToolStripMenuItem
			// 
			this.printToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("printToolStripMenuItem.Image")));
			this.printToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.printToolStripMenuItem.Name = "printToolStripMenuItem";
			this.printToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.P)));
			this.printToolStripMenuItem.Size = new System.Drawing.Size(139, 22);
			this.printToolStripMenuItem.Text = "&Print";
			// 
			// printPreviewToolStripMenuItem
			// 
			this.printPreviewToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("printPreviewToolStripMenuItem.Image")));
			this.printPreviewToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.printPreviewToolStripMenuItem.Name = "printPreviewToolStripMenuItem";
			this.printPreviewToolStripMenuItem.Size = new System.Drawing.Size(139, 22);
			this.printPreviewToolStripMenuItem.Text = "Print Pre&view";
			// 
			// toolStripSeparator2
			// 
			this.toolStripSeparator2.Name = "toolStripSeparator2";
			this.toolStripSeparator2.Size = new System.Drawing.Size(136, 6);
			// 
			// exitToolStripMenuItem
			// 
			this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
			this.exitToolStripMenuItem.Size = new System.Drawing.Size(139, 22);
			this.exitToolStripMenuItem.Text = "E&xit";
			// 
			// editToolStripMenuItem
			// 
			this.editToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.undoToolStripMenuItem,
            this.redoToolStripMenuItem,
            this.toolStripSeparator3,
            this.cutToolStripMenuItem,
            this.copyToolStripMenuItem,
            this.pasteToolStripMenuItem,
            this.toolStripSeparator4,
            this.selectAllToolStripMenuItem});
			this.editToolStripMenuItem.Name = "editToolStripMenuItem";
			this.editToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
			this.editToolStripMenuItem.Text = "&Edit";
			// 
			// undoToolStripMenuItem
			// 
			this.undoToolStripMenuItem.Name = "undoToolStripMenuItem";
			this.undoToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Z)));
			this.undoToolStripMenuItem.Size = new System.Drawing.Size(139, 22);
			this.undoToolStripMenuItem.Text = "&Undo";
			// 
			// redoToolStripMenuItem
			// 
			this.redoToolStripMenuItem.Name = "redoToolStripMenuItem";
			this.redoToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Y)));
			this.redoToolStripMenuItem.Size = new System.Drawing.Size(139, 22);
			this.redoToolStripMenuItem.Text = "&Redo";
			// 
			// toolStripSeparator3
			// 
			this.toolStripSeparator3.Name = "toolStripSeparator3";
			this.toolStripSeparator3.Size = new System.Drawing.Size(136, 6);
			// 
			// cutToolStripMenuItem
			// 
			this.cutToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("cutToolStripMenuItem.Image")));
			this.cutToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.cutToolStripMenuItem.Name = "cutToolStripMenuItem";
			this.cutToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.X)));
			this.cutToolStripMenuItem.Size = new System.Drawing.Size(139, 22);
			this.cutToolStripMenuItem.Text = "Cu&t";
			// 
			// copyToolStripMenuItem
			// 
			this.copyToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("copyToolStripMenuItem.Image")));
			this.copyToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.copyToolStripMenuItem.Name = "copyToolStripMenuItem";
			this.copyToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.C)));
			this.copyToolStripMenuItem.Size = new System.Drawing.Size(139, 22);
			this.copyToolStripMenuItem.Text = "&Copy";
			// 
			// pasteToolStripMenuItem
			// 
			this.pasteToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("pasteToolStripMenuItem.Image")));
			this.pasteToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.pasteToolStripMenuItem.Name = "pasteToolStripMenuItem";
			this.pasteToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.V)));
			this.pasteToolStripMenuItem.Size = new System.Drawing.Size(139, 22);
			this.pasteToolStripMenuItem.Text = "&Paste";
			// 
			// toolStripSeparator4
			// 
			this.toolStripSeparator4.Name = "toolStripSeparator4";
			this.toolStripSeparator4.Size = new System.Drawing.Size(136, 6);
			// 
			// selectAllToolStripMenuItem
			// 
			this.selectAllToolStripMenuItem.Name = "selectAllToolStripMenuItem";
			this.selectAllToolStripMenuItem.Size = new System.Drawing.Size(139, 22);
			this.selectAllToolStripMenuItem.Text = "Select &All";
			// 
			// toolsToolStripMenuItem
			// 
			this.toolsToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.customizeToolStripMenuItem,
            this.optionsToolStripMenuItem});
			this.toolsToolStripMenuItem.Name = "toolsToolStripMenuItem";
			this.toolsToolStripMenuItem.Size = new System.Drawing.Size(45, 20);
			this.toolsToolStripMenuItem.Text = "&Tools";
			// 
			// customizeToolStripMenuItem
			// 
			this.customizeToolStripMenuItem.Name = "customizeToolStripMenuItem";
			this.customizeToolStripMenuItem.Size = new System.Drawing.Size(125, 22);
			this.customizeToolStripMenuItem.Text = "&Customize";
			// 
			// optionsToolStripMenuItem
			// 
			this.optionsToolStripMenuItem.Name = "optionsToolStripMenuItem";
			this.optionsToolStripMenuItem.Size = new System.Drawing.Size(125, 22);
			this.optionsToolStripMenuItem.Text = "&Options";
			// 
			// helpToolStripMenuItem
			// 
			this.helpToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.contentsToolStripMenuItem,
            this.indexToolStripMenuItem,
            this.searchToolStripMenuItem,
            this.toolStripSeparator5,
            this.aboutToolStripMenuItem});
			this.helpToolStripMenuItem.Name = "helpToolStripMenuItem";
			this.helpToolStripMenuItem.Size = new System.Drawing.Size(41, 20);
			this.helpToolStripMenuItem.Text = "&Help";
			// 
			// contentsToolStripMenuItem
			// 
			this.contentsToolStripMenuItem.Name = "contentsToolStripMenuItem";
			this.contentsToolStripMenuItem.Size = new System.Drawing.Size(119, 22);
			this.contentsToolStripMenuItem.Text = "&Contents";
			// 
			// indexToolStripMenuItem
			// 
			this.indexToolStripMenuItem.Name = "indexToolStripMenuItem";
			this.indexToolStripMenuItem.Size = new System.Drawing.Size(119, 22);
			this.indexToolStripMenuItem.Text = "&Index";
			// 
			// searchToolStripMenuItem
			// 
			this.searchToolStripMenuItem.Name = "searchToolStripMenuItem";
			this.searchToolStripMenuItem.Size = new System.Drawing.Size(119, 22);
			this.searchToolStripMenuItem.Text = "&Search";
			// 
			// toolStripSeparator5
			// 
			this.toolStripSeparator5.Name = "toolStripSeparator5";
			this.toolStripSeparator5.Size = new System.Drawing.Size(116, 6);
			// 
			// aboutToolStripMenuItem
			// 
			this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
			this.aboutToolStripMenuItem.Size = new System.Drawing.Size(119, 22);
			this.aboutToolStripMenuItem.Text = "&About...";
			this.aboutToolStripMenuItem.Click += new System.EventHandler(this.aboutToolStripMenuItem_Click);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 71);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(125, 13);
			this.label2.TabIndex = 25;
			this.label2.Text = "NASA STS TV Schedule";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(11, 244);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(110, 13);
			this.label3.TabIndex = 0;
			this.label3.Text = "Begin Date of Mission";
			// 
			// dtpOutlook
			// 
			this.dtpOutlook.AccessibleDescription = "Change the beginning date to look for appointment items in Outlook";
			this.dtpOutlook.AccessibleName = "Outlook Begin Search Date";
			this.dtpOutlook.Location = new System.Drawing.Point(124, 240);
			this.dtpOutlook.MaxDate = new System.DateTime(2010, 12, 31, 0, 0, 0, 0);
			this.dtpOutlook.MinDate = new System.DateTime(1981, 11, 12, 0, 0, 0, 0);
			this.dtpOutlook.Name = "dtpOutlook";
			this.dtpOutlook.Size = new System.Drawing.Size(200, 20);
			this.dtpOutlook.TabIndex = 2;
			this.dtpOutlook.Value = new System.DateTime(2007, 10, 16, 0, 0, 0, 0);
			this.dtpOutlook.MouseLeave += new System.EventHandler(this.dtpOutlook_MouseLeave);
			this.dtpOutlook.Leave += new System.EventHandler(this.dtpOutlook_Leave);
			this.dtpOutlook.CloseUp += new System.EventHandler(this.dtpOutlook_CloseUp);
			this.dtpOutlook.MouseHover += new System.EventHandler(this.dtpOutlook_MouseHover);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(11, 266);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(97, 13);
			this.label4.TabIndex = 26;
			this.label4.Text = "Outlook Categories";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(330, 244);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(98, 13);
			this.label5.TabIndex = 27;
			this.label5.Text = "Viewing Time Zone";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 332);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(92, 13);
			this.label6.TabIndex = 28;
			this.label6.Text = "Outlook Schedule";
			// 
			// MainForm
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
			this.ClientSize = new System.Drawing.Size(795, 573);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.dtpOutlook);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.chklbOutlookCategories);
			this.Controls.Add(this.btnBulkImport);
			this.Controls.Add(this.btnSmartSelect);
			this.Controls.Add(this.lblNasaStsTVScheduleFile);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.statusStrip);
			this.Controls.Add(this.menuStrip);
			this.Controls.Add(this.btnRefreshOutlookCategories);
			this.Controls.Add(this.cmbxTimeZones);
			this.Controls.Add(this.btnUnselectAllOutlook);
			this.Controls.Add(this.btnUnselectAllExcel);
			this.Controls.Add(this.btnSelectAllOutlook);
			this.Controls.Add(this.btnSelectAllExcel);
			this.Controls.Add(this.chkOrbit);
			this.Controls.Add(this.chkDocked);
			this.Controls.Add(this.chkInteropExcel);
			this.Controls.Add(this.btnExitApplication);
			this.Controls.Add(this.btnTransferTVSchedule);
			this.Controls.Add(this.btnRemoveMarkedEntries);
			this.Controls.Add(this.dgvOutlook);
			this.Controls.Add(this.dgvExcelSchedule);
			this.Controls.Add(this.btnOpenNasaTvSchedule);
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.MainMenuStrip = this.menuStrip;
			this.Name = "MainForm";
			this.Text = "NASA Space Shuttle TV Schedule Transfer to Outlook Calendar";
			this.Load += new System.EventHandler(this.MainForm_Load);
			((System.ComponentModel.ISupportInitialize)(this.dgvExcelSchedule)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dgvOutlook)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnOpenNasaTvSchedule;
		private System.Windows.Forms.DataGridView dgvExcelSchedule;
        private System.Windows.Forms.DataGridView dgvOutlook;
        private System.Windows.Forms.Button btnRemoveMarkedEntries;
        private System.Windows.Forms.Button btnTransferTVSchedule;
        private System.Windows.Forms.Button btnExitApplication;
        private System.Windows.Forms.OpenFileDialog ofExcelSchedule;
        private System.Windows.Forms.CheckBox chkInteropExcel;
        private System.Windows.Forms.CheckBox chkDocked;
        private System.Windows.Forms.CheckBox chkOrbit;
        private System.Windows.Forms.CheckedListBox chklbOutlookCategories;
        private System.Windows.Forms.Button btnSelectAllExcel;
		private System.Windows.Forms.Button btnSelectAllOutlook;
        private System.Windows.Forms.Button btnUnselectAllExcel;
        private System.Windows.Forms.Button btnUnselectAllOutlook;
        private System.Windows.Forms.ComboBox cmbxTimeZones;
		private System.Windows.Forms.Button btnRefreshOutlookCategories;
		private System.Windows.Forms.StatusStrip statusStrip;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel;
		private System.Windows.Forms.ToolStripProgressBar toolStripProgressBar;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label lblNasaStsTVScheduleFile;
		private System.Windows.Forms.Button btnSmartSelect;
		private System.Windows.Forms.DataGridViewCheckBoxColumn REMOVE_OL;
		private System.Windows.Forms.DataGridViewTextBoxColumn BEGIN_DATE_OL;
		private System.Windows.Forms.DataGridViewTextBoxColumn END_DATE_OL;
		private System.Windows.Forms.DataGridViewTextBoxColumn SUBJECT_OL;
		private System.Windows.Forms.DataGridViewTextBoxColumn SITE_OL;
		private System.Windows.Forms.Button btnBulkImport;
        private System.Windows.Forms.DataGridViewCheckBoxColumn ID_ADD;
        private System.Windows.Forms.DataGridViewTextBoxColumn ORBIT_TV;
        private System.Windows.Forms.DataGridViewTextBoxColumn BEGIN_DATE_TV;
        private System.Windows.Forms.DataGridViewTextBoxColumn END_DATE_TV;
        private System.Windows.Forms.DataGridViewCheckBoxColumn REMINDER_TV;
        private System.Windows.Forms.DataGridViewCheckBoxColumn CHANGED_TV;
        private System.Windows.Forms.DataGridViewTextBoxColumn SUBJECT_TV;
        private System.Windows.Forms.DataGridViewTextBoxColumn SITE_TV;
		private System.Windows.Forms.MenuStrip menuStrip;
		private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem newToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem openToolStripMenuItem;
		private System.Windows.Forms.ToolStripSeparator toolStripSeparator;
		private System.Windows.Forms.ToolStripMenuItem saveToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem saveAsToolStripMenuItem;
		private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
		private System.Windows.Forms.ToolStripMenuItem printToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem printPreviewToolStripMenuItem;
		private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
		private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem editToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem undoToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem redoToolStripMenuItem;
		private System.Windows.Forms.ToolStripSeparator toolStripSeparator3;
		private System.Windows.Forms.ToolStripMenuItem cutToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem copyToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem pasteToolStripMenuItem;
		private System.Windows.Forms.ToolStripSeparator toolStripSeparator4;
		private System.Windows.Forms.ToolStripMenuItem selectAllToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem toolsToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem customizeToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem optionsToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem helpToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem contentsToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem indexToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem searchToolStripMenuItem;
		private System.Windows.Forms.ToolStripSeparator toolStripSeparator5;
		private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.DateTimePicker dtpOutlook;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
    }
}

