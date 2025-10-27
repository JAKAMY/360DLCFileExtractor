// Type: wxPirs.MainForm
// Assembly: wxPirs, Version=1.1.0.42, Culture=neutral, PublicKeyToken=null
// Assembly location: C:\Users\jamesc\Downloads\wxPirs-1.1\wxPirs.exe

using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using wxPirs.Properties;
using WxTools;
using System.Collections.Generic;
using System.Text;

namespace wxPirs
{
    public partial class MainForm : Form
    {
        private IContainer components;
        private MenuStrip menuStrip;
        private StatusStrip statusStrip;
        private ToolStripContainer toolStripContainer1;
        private SplitContainer splitContainerH;
        private SplitContainer splitContainerV;
        private TreeView treeView;
        private ListView listView;
        private ColumnHeader columnHeaderName;
        private TextBox textBoxLog;
        private ImageList imageListTreeView;
        private ToolStripMenuItem fileToolStripMenuItem;
        private ToolStripMenuItem exitToolStripMenuItem;
        private ToolStrip fileToolStrip;
        private ToolStripButton openFileToolStripButton;
        private OpenFileDialog openFileDialog;
        private ColumnHeader columnHeaderSize;
        private ColumnHeader columnHeaderDateModified;
        private ImageList imageList;
        private ColumnHeader columnHeaderCluster;
        private ContextMenuStrip contextMenuStrip;
        private ToolStripMenuItem extractFileToolStripMenuItem;
        private ContextMenuStrip contextMenuStripMulti;
        private ToolStripMenuItem extractFilesToolStripMenuItem;
        private SaveFileDialog saveFileDialog;
        private ColumnHeader columnHeaderStatus;
        private ToolStripMenuItem toolStripMenuItem1;
        private ToolStripMenuItem aboutPirsToolStripMenuItem;
        private ToolStripStatusLabel toolStripStatusLabelVersion;
        private ToolStripMenuItem openFileToolStripMenuItem;
        private ToolStripSeparator toolStripMenuItem2;
        private ToolStripStatusLabel toolStripStatusLabelSeparator;
        private ToolStripStatusLabel toolStripStatusLabelGael360;
        private FolderBrowserDialog folderBrowserDialog;
        private ContextMenuStrip contextMenuStripFolder;
        private ToolStripMenuItem extractFolderToolStripMenuItem;
        private ToolStripButton extractAllToolStripButton;
        private ToolStripMenuItem extractAllToolStripMenuItem;

        // New controls
        private ToolStripButton openFolderToolStripButton;
        private ToolStripMenuItem openFolderToolStripMenuItem;
        private SplitContainer splitContainerList; // hosts listView and queue listbox
        private ListBox listBoxQueue;

        // New: out folder, process queue and queue context menu
        private ToolStripButton setOutFolderToolStripButton;
        private ToolStripButton batchToolStripButton;
        private ToolStripStatusLabel toolStripStatusLabelOutFolder;
        private ContextMenuStrip contextMenuQueue;
        private ToolStripMenuItem removeQueueToolStripMenuItem;
        private string outFolderPath = "";

        // Added for new features
        private ToolStripButton stopBatchToolStripButton;
        private ToolStripProgressBar toolStripProgressBar;
        private volatile bool stopBatchRequested = false;
        private bool isBatchRunning = false;
        private string batchLogFilePath = null;

        // New UI elements for layout requested
        private ToolStripSeparator toolStripSeparator1;
        private ToolStripSeparator toolStripSeparator2;
        private ToolStripLabel toolStripLabelBatch;

        // queue statuses
        private enum QueueItemStatus { Pending, Success, Error }
        private List<QueueItemStatus> queueStatuses = new List<QueueItemStatus>();

        public static int MAGIC_PIRS;
        public static int MAGIC_LIVE;
        public static int MAGIC_CON_;
        public static long PIRS_TYPE1;
        public static long PIRS_TYPE2;
        public static long PIRS_BASE;
        private string[] args;
        private WxReader wr;
        private FileStream fs;
        private BinaryReader br;
        private long pirs_offset;
        private long pirs_start;

        static MainForm()
        {
            MainForm.MAGIC_PIRS = 1346982483;
            MainForm.MAGIC_LIVE = 1279874629;
            MainForm.MAGIC_CON_ = 1129270816;
            MainForm.PIRS_TYPE1 = 4096L;
            MainForm.PIRS_TYPE2 = 8192L;
            MainForm.PIRS_BASE = 45056L;
        }

        public MainForm(string[] args)
        {
            this.wr = new WxReader();
            this.args = args;
            this.InitializeComponent();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && this.components != null)
                this.components.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.components = new Container();
            ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(MainForm));

            // Instantiate controls (each only once)
            this.menuStrip = new MenuStrip();
            this.fileToolStripMenuItem = new ToolStripMenuItem();
            this.openFileToolStripMenuItem = new ToolStripMenuItem();
            this.openFolderToolStripMenuItem = new ToolStripMenuItem();
            this.extractAllToolStripMenuItem = new ToolStripMenuItem();
            this.toolStripMenuItem2 = new ToolStripSeparator();
            this.exitToolStripMenuItem = new ToolStripMenuItem();
            this.toolStripMenuItem1 = new ToolStripMenuItem();
            this.aboutPirsToolStripMenuItem = new ToolStripMenuItem();

            this.statusStrip = new StatusStrip();
            this.toolStripStatusLabelVersion = new ToolStripStatusLabel();
            this.toolStripStatusLabelSeparator = new ToolStripStatusLabel();
            this.toolStripStatusLabelGael360 = new ToolStripStatusLabel();
            this.toolStripStatusLabelOutFolder = new ToolStripStatusLabel();

            this.toolStripContainer1 = new ToolStripContainer();
            this.splitContainerH = new SplitContainer();
            this.splitContainerV = new SplitContainer();

            this.treeView = new TreeView();
            this.imageListTreeView = new ImageList(this.components);

            // split container between list and queue
            this.splitContainerList = new SplitContainer();
            this.listView = new ListView();
            this.columnHeaderName = new ColumnHeader();
            this.columnHeaderSize = new ColumnHeader();
            this.columnHeaderCluster = new ColumnHeader();
            this.columnHeaderDateModified = new ColumnHeader();
            this.columnHeaderStatus = new ColumnHeader();
            this.listBoxQueue = new ListBox();

            this.contextMenuQueue = new ContextMenuStrip(this.components);
            this.removeQueueToolStripMenuItem = new ToolStripMenuItem();

            this.imageList = new ImageList(this.components);
            this.textBoxLog = new TextBox();

            this.fileToolStrip = new ToolStrip();
            this.openFileToolStripButton = new ToolStripButton();
            this.openFolderToolStripButton = new ToolStripButton();
            this.setOutFolderToolStripButton = new ToolStripButton();
            this.batchToolStripButton = new ToolStripButton();
            this.extractAllToolStripButton = new ToolStripButton();

            // New stop button and progress bar
            this.stopBatchToolStripButton = new ToolStripButton();
            this.toolStripProgressBar = new ToolStripProgressBar();

            // New separators and label
            this.toolStripSeparator1 = new ToolStripSeparator();
            this.toolStripSeparator2 = new ToolStripSeparator();
            this.toolStripLabelBatch = new ToolStripLabel();

            this.contextMenuStripFolder = new ContextMenuStrip(this.components);
            this.extractFolderToolStripMenuItem = new ToolStripMenuItem();

            this.openFileDialog = new OpenFileDialog();
            this.contextMenuStrip = new ContextMenuStrip(this.components);
            this.extractFileToolStripMenuItem = new ToolStripMenuItem();
            this.contextMenuStripMulti = new ContextMenuStrip(this.components);
            this.extractFilesToolStripMenuItem = new ToolStripMenuItem();
            this.saveFileDialog = new SaveFileDialog();
            this.folderBrowserDialog = new FolderBrowserDialog();

            // Suspend layouts
            this.menuStrip.SuspendLayout();
            this.statusStrip.SuspendLayout();
            this.toolStripContainer1.ContentPanel.SuspendLayout();
            this.toolStripContainer1.TopToolStripPanel.SuspendLayout();
            this.toolStripContainer1.SuspendLayout();
            this.splitContainerH.Panel1.SuspendLayout();
            this.splitContainerH.Panel2.SuspendLayout();
            this.splitContainerH.SuspendLayout();
            this.splitContainerV.Panel1.SuspendLayout();
            this.splitContainerV.Panel2.SuspendLayout();
            this.splitContainerV.ResumeLayout();
            this.splitContainerList.Panel1.SuspendLayout();
            this.splitContainerList.Panel2.SuspendLayout();
            this.splitContainerList.ResumeLayout();
            this.fileToolStrip.SuspendLayout();
            this.contextMenuStripFolder.SuspendLayout();
            this.contextMenuStrip.SuspendLayout();
            this.contextMenuStripMulti.SuspendLayout();
            this.contextMenuQueue.SuspendLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

            // MenuStrip
            this.menuStrip.Items.AddRange(new ToolStripItem[] { this.fileToolStripMenuItem, this.toolStripMenuItem1 });
            this.menuStrip.Location = new Point(0, 0);
            this.menuStrip.Name = "menuStrip";
            this.menuStrip.Size = new Size(729, 24);
            this.menuStrip.TabIndex = 0;
            this.menuStrip.Text = "menuStrip1";

            // File menu and items
            this.fileToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] {
 this.openFileToolStripMenuItem,
 this.openFolderToolStripMenuItem,
 this.extractAllToolStripMenuItem,
 this.toolStripMenuItem2,
 this.exitToolStripMenuItem
});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new Size(35, 20);
            this.fileToolStripMenuItem.Text = "File";

            this.openFileToolStripMenuItem.Image = (Image)Resources.OpenFile;
            this.openFileToolStripMenuItem.Name = "openFileToolStripMenuItem";
            this.openFileToolStripMenuItem.Size = new Size(134, 22);
            this.openFileToolStripMenuItem.Text = "Open...";
            this.openFileToolStripMenuItem.Click += openFile;

            this.openFolderToolStripMenuItem.Image = (Image)Resources.OpenFolder;
            this.openFolderToolStripMenuItem.Name = "openFolderToolStripMenuItem";
            this.openFolderToolStripMenuItem.Size = new Size(134, 22);
            this.openFolderToolStripMenuItem.Text = "Open folder...";
            this.openFolderToolStripMenuItem.Click += new EventHandler(this.openFolderMenu_Click);

            this.extractAllToolStripMenuItem.Image = (Image)Resources.Extract;
            this.extractAllToolStripMenuItem.Name = "extractAllToolStripMenuItem";
            this.extractAllToolStripMenuItem.Size = new Size(134, 22);
            this.extractAllToolStripMenuItem.Text = "Extract all...";
            this.extractAllToolStripMenuItem.Click += extractAll;

            this.toolStripMenuItem2.Name = "toolStripMenuItem2";
            this.toolStripMenuItem2.Size = new Size(131, 6);

            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new Size(134, 22);
            this.exitToolStripMenuItem.Text = "Exit";
            this.exitToolStripMenuItem.Click += exitToolStripMenuItem_Click;

            this.toolStripMenuItem1.DropDownItems.AddRange(new ToolStripItem[] { this.aboutPirsToolStripMenuItem });
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new Size(24, 20);
            this.toolStripMenuItem1.Text = "?";

            this.aboutPirsToolStripMenuItem.Name = "aboutPirsToolStripMenuItem";
            this.aboutPirsToolStripMenuItem.Size = new Size(123, 22);
            this.aboutPirsToolStripMenuItem.Text = "About Pirs";
            this.aboutPirsToolStripMenuItem.Click += aboutPirsToolStripMenuItem_Click;

            // StatusStrip (include OutFolder label)
            this.statusStrip.Items.AddRange(new ToolStripItem[] {
 this.toolStripStatusLabelVersion,
 this.toolStripStatusLabelSeparator,
 this.toolStripStatusLabelGael360,
 this.toolStripStatusLabelOutFolder
});
            this.statusStrip.Location = new Point(0, 373);
            this.statusStrip.Name = "statusStrip";
            this.statusStrip.Size = new Size(729, 22);
            this.statusStrip.TabIndex = 1;
            this.statusStrip.Text = "statusStrip1";

            this.toolStripStatusLabelVersion.Name = "toolStripStatusLabelVersion";
            this.toolStripStatusLabelVersion.Size = new Size(664, 17);
            this.toolStripStatusLabelVersion.Spring = true;
            this.toolStripStatusLabelVersion.Text = "wxPirs";
            this.toolStripStatusLabelVersion.TextAlign = ContentAlignment.MiddleLeft;

            this.toolStripStatusLabelSeparator.BorderSides = ToolStripStatusLabelBorderSides.All;
            this.toolStripStatusLabelSeparator.BorderStyle = Border3DStyle.Sunken;
            this.toolStripStatusLabelSeparator.Name = "toolStripStatusLabelSeparator";
            this.toolStripStatusLabelSeparator.Size = new Size(4, 17);

            this.toolStripStatusLabelGael360.Name = "toolStripStatusLabelGael360";
            this.toolStripStatusLabelGael360.Size = new Size(46, 17);
            this.toolStripStatusLabelGael360.Text = "Gael360";

            this.toolStripStatusLabelOutFolder.Name = "toolStripStatusLabelOutFolder";
            this.toolStripStatusLabelOutFolder.Size = new Size(200, 17);
            this.toolStripStatusLabelOutFolder.Text = "Out Folder: (not set)";

            // ToolStripContainer
            this.toolStripContainer1.ContentPanel.Controls.Add(this.splitContainerH);
            this.toolStripContainer1.ContentPanel.Size = new Size(729, 324);
            this.toolStripContainer1.Dock = DockStyle.Fill;
            this.toolStripContainer1.Location = new Point(0, 24);
            this.toolStripContainer1.Name = "toolStripContainer1";
            this.toolStripContainer1.Size = new Size(729, 349);
            this.toolStripContainer1.TabIndex = 2;
            this.toolStripContainer1.Text = "toolStripContainer1";
            this.toolStripContainer1.TopToolStripPanel.Controls.Add(this.fileToolStrip);

            // Top/bottom split
            this.splitContainerH.Dock = DockStyle.Fill;
            this.splitContainerH.FixedPanel = FixedPanel.Panel2;
            this.splitContainerH.Location = new Point(0, 0);
            this.splitContainerH.Name = "splitContainerH";
            this.splitContainerH.Orientation = Orientation.Horizontal;
            this.splitContainerH.Panel1.Controls.Add(this.splitContainerV);
            this.splitContainerH.Panel2.Controls.Add(this.textBoxLog);
            this.splitContainerH.Size = new Size(729, 324);
            this.splitContainerH.SplitterDistance = 243;
            this.splitContainerH.TabIndex = 0;

            // left/right split
            this.splitContainerV.Dock = DockStyle.Fill;
            this.splitContainerV.FixedPanel = FixedPanel.Panel1;
            this.splitContainerV.Location = new Point(0, 0);
            this.splitContainerV.Name = "splitContainerV";
            this.splitContainerV.Panel1.Controls.Add(this.treeView);
            this.splitContainerV.Panel2.Controls.Add(this.splitContainerList);
            this.splitContainerV.Size = new Size(729, 243);
            this.splitContainerV.SplitterDistance = 180;
            this.splitContainerV.TabIndex = 0;

            // treeView
            this.treeView.Dock = DockStyle.Fill;
            this.treeView.ImageIndex = 1;
            this.treeView.ImageList = this.imageListTreeView;
            this.treeView.Location = new Point(0, 0);
            this.treeView.Name = "treeView";
            this.treeView.SelectedImageIndex = 2;
            this.treeView.ShowRootLines = false;
            this.treeView.Size = new Size(180, 243);
            this.treeView.TabIndex = 0;
            this.treeView.AfterSelect += treeView_AfterSelect;
            this.treeView.NodeMouseClick += treeView_NodeMouseClick;

            // this.imageListTreeView.ImageStream = (ImageListStreamer) componentResourceManager.GetObject("imageListTreeView.ImageStream");
            this.imageListTreeView.TransparentColor = Color.Fuchsia;

            // splitContainerList hosts listView and queue
            this.splitContainerList.Dock = DockStyle.Fill;
            this.splitContainerList.Location = new Point(0, 0);
            this.splitContainerList.Name = "splitContainerList";
            this.splitContainerList.Panel1.Controls.Add(this.listView);
            this.splitContainerList.Panel2.Controls.Add(this.listBoxQueue);
            this.splitContainerList.Size = new Size(545, 243);
            this.splitContainerList.SplitterDistance = 380;
            this.splitContainerList.TabIndex = 0;

            // listView and its columns
            this.listView.Columns.AddRange(new ColumnHeader[] {
 this.columnHeaderName,
 this.columnHeaderSize,
 this.columnHeaderCluster,
 this.columnHeaderDateModified,
 this.columnHeaderStatus
});
            this.listView.Dock = DockStyle.Fill;
            this.listView.FullRowSelect = true;
            this.listView.Location = new Point(0, 0);
            this.listView.Name = "listView";
            this.listView.ShowItemToolTips = true;
            this.listView.Size = new Size(380, 243);
            this.listView.SmallImageList = this.imageList;
            this.listView.TabIndex = 0;
            this.listView.UseCompatibleStateImageBehavior = false;
            this.listView.View = View.Details;
            this.listView.MouseClick += listView_MouseClick;

            this.columnHeaderName.Text = "Name";
            this.columnHeaderName.Width = 200;
            this.columnHeaderSize.Text = "Size";
            this.columnHeaderSize.TextAlign = HorizontalAlignment.Right;
            this.columnHeaderCluster.Text = "Cluster";
            this.columnHeaderCluster.TextAlign = HorizontalAlignment.Right;
            this.columnHeaderDateModified.Text = "Date modified";
            this.columnHeaderDateModified.Width = 120;
            this.columnHeaderStatus.Text = "Status";
            this.columnHeaderStatus.Width = 80;

            // queue list box - owner draw to color items per status
            this.listBoxQueue.Dock = DockStyle.Fill;
            this.listBoxQueue.FormattingEnabled = true;
            this.listBoxQueue.Location = new Point(0, 0);
            this.listBoxQueue.Name = "listBoxQueue";
            this.listBoxQueue.Size = new Size(161, 238);
            this.listBoxQueue.TabIndex = 0;
            this.listBoxQueue.DoubleClick += new EventHandler(this.listBoxQueue_DoubleClick);
            this.listBoxQueue.MouseUp += new MouseEventHandler(this.listBoxQueue_MouseUp);

            // Use owner draw to color rows
            this.listBoxQueue.DrawMode = DrawMode.OwnerDrawFixed;
            this.listBoxQueue.DrawItem += listBoxQueue_DrawItem;

            // context menu for queue
            this.contextMenuQueue.Items.AddRange(new ToolStripItem[] { this.removeQueueToolStripMenuItem });
            this.removeQueueToolStripMenuItem.Name = "removeQueueToolStripMenuItem";
            this.removeQueueToolStripMenuItem.Text = "Remove";
            this.removeQueueToolStripMenuItem.Click += new EventHandler(this.removeQueueToolStripMenuItem_Click);

            // TextBox log
            this.textBoxLog.Dock = DockStyle.Fill;
            this.textBoxLog.Location = new Point(0, 0);
            this.textBoxLog.Multiline = true;
            this.textBoxLog.Name = "textBoxLog";
            this.textBoxLog.ScrollBars = ScrollBars.Both;
            this.textBoxLog.Size = new Size(729, 77);
            this.textBoxLog.TabIndex = 0;

            // ToolStrip (top) - arranged per your requested order
            //
            // Order:
            // 1) Open file (icon + text)
            // 2) Extract files (icon + text)
            // 3) Separator
            // 4) Label "Batch" (bold)
            // 5) Open Batch folder (icon + text)
            // 6) Open output folder (icon + text)
            // 7) Process queue (icon + text)
            // 8) Separator
            // 9) Stop (icon/text)
            // 10) Progress bar
            //
            this.fileToolStrip.Dock = DockStyle.None;
            this.fileToolStrip.Items.Clear();
            // configure buttons to show ImageAndText where requested
            this.openFileToolStripButton.DisplayStyle = ToolStripItemDisplayStyle.ImageAndText;
            this.openFileToolStripButton.Image = (Image)Resources.OpenFile;
            this.openFileToolStripButton.ImageTransparentColor = Color.Magenta;
            this.openFileToolStripButton.Name = "openFileToolStripButton";
            this.openFileToolStripButton.Size = new Size(90, 22);
            this.openFileToolStripButton.Text = "Open file";
            this.openFileToolStripButton.Click += openFile;

            this.extractAllToolStripButton.DisplayStyle = ToolStripItemDisplayStyle.ImageAndText;
            this.extractAllToolStripButton.Image = (Image)Resources.Extract;
            this.extractAllToolStripButton.ImageTransparentColor = Color.Magenta;
            this.extractAllToolStripButton.Name = "extractAllToolStripButton";
            this.extractAllToolStripButton.Size = new Size(110, 22);
            this.extractAllToolStripButton.Text = "Extract files";
            // reuse existing handler if appropriate; original extractAll used for "Extract all"
            // Keep Click wired to extractAll for compatibility
            this.extractAllToolStripButton.Click += extractAll;

            // separators and label
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new Size(6, 25);

            this.toolStripLabelBatch.Name = "toolStripLabelBatch";
            this.toolStripLabelBatch.Size = new Size(40, 22);
            // set bold font for label
            this.toolStripLabelBatch.Font = new Font(this.toolStripLabelBatch.Font, FontStyle.Bold);
            this.toolStripLabelBatch.Text = "Batch";

            // open Batch folder button (icon + text)
            this.openFolderToolStripButton.DisplayStyle = ToolStripItemDisplayStyle.ImageAndText;
            this.openFolderToolStripButton.Image = (Image)wxPirs.Properties.Resources.OpenFolder;
            this.openFolderToolStripButton.ImageTransparentColor = Color.Magenta;
            this.openFolderToolStripButton.Name = "openFolderToolStripButton";
            this.openFolderToolStripButton.Size = new Size(100, 22);
            this.openFolderToolStripButton.Text = "Open Batch folder";
            this.openFolderToolStripButton.Click += new EventHandler(this.openFolderMenu_Click);

            // open/set output folder (icon + text)
            this.setOutFolderToolStripButton.DisplayStyle = ToolStripItemDisplayStyle.ImageAndText;
            this.setOutFolderToolStripButton.Image = (Image)Resources.OpenOutFolder;
            this.setOutFolderToolStripButton.ImageTransparentColor = Color.Magenta;
            this.setOutFolderToolStripButton.Name = "setOutFolderToolStripButton";
            this.setOutFolderToolStripButton.Size = new Size(130, 22);
            this.setOutFolderToolStripButton.Text = "Open output folder";
            this.setOutFolderToolStripButton.Click += new EventHandler(this.setOutFolderToolStripButton_Click);

            // Process queue (was batchToolStripButton) (icon + text)
            this.batchToolStripButton.DisplayStyle = ToolStripItemDisplayStyle.ImageAndText;
            this.batchToolStripButton.Image = (Image)Resources.Batch;
            this.batchToolStripButton.ImageTransparentColor = Color.Magenta;
            this.batchToolStripButton.Name = "batchToolStripButton";
            this.batchToolStripButton.Size = new Size(120, 22);
            this.batchToolStripButton.Text = "Process queue";
            this.batchToolStripButton.Click += new EventHandler(this.batchToolStripButton_Click);

            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new Size(6, 25);

            // Stop button (icon + text)
            this.stopBatchToolStripButton.DisplayStyle = ToolStripItemDisplayStyle.ImageAndText;
            this.stopBatchToolStripButton.Image = (Image)Resources.Stop;
            this.stopBatchToolStripButton.ImageTransparentColor = Color.Magenta;
            this.stopBatchToolStripButton.Name = "stopBatchToolStripButton";
            this.stopBatchToolStripButton.Size = new Size(60, 22);
            this.stopBatchToolStripButton.Text = "Stop";
            this.stopBatchToolStripButton.Enabled = false;
            this.stopBatchToolStripButton.Click += new EventHandler(this.stopBatchToolStripButton_Click);

            // Progress bar
            this.toolStripProgressBar.Name = "toolStripProgressBar";
            this.toolStripProgressBar.Size = new Size(120, 16);
            this.toolStripProgressBar.Visible = false;

            // Add items to toolstrip in required order
            this.fileToolStrip.Items.AddRange(new ToolStripItem[] {
   this.openFileToolStripButton,
   this.extractAllToolStripButton,
   this.toolStripSeparator1,
   this.toolStripLabelBatch,
   this.openFolderToolStripButton,
   this.setOutFolderToolStripButton,
   this.batchToolStripButton,
   this.toolStripSeparator2,
   this.stopBatchToolStripButton,
   this.toolStripProgressBar
});

            this.fileToolStrip.Location = new Point(3, 0);
            this.fileToolStrip.Name = "fileToolStrip";
            this.fileToolStrip.Size = new Size(400, 25);
            this.fileToolStrip.TabIndex = 0;
            this.fileToolStrip.Text = "File";

            // Context menu for folders/files
            this.contextMenuStripFolder.Items.AddRange(new ToolStripItem[] { this.extractFolderToolStripMenuItem });
            this.contextMenuStripFolder.Name = "contextMenuStripFolder";
            this.contextMenuStripFolder.Size = new Size(141, 26);
            this.extractFolderToolStripMenuItem.Name = "extractFolderToolStripMenuItem";
            this.extractFolderToolStripMenuItem.Size = new Size(140, 22);
            this.extractFolderToolStripMenuItem.Text = "Extract folder";
            this.extractFolderToolStripMenuItem.Click += extractFolder;

            this.openFileDialog.Filter = "All files (*.*)|*.*";

            this.contextMenuStrip.Items.AddRange(new ToolStripItem[] { this.extractFileToolStripMenuItem });
            this.contextMenuStrip.Name = "contextMenuStrip";
            this.contextMenuStrip.Size = new Size(127, 26);
            this.extractFileToolStripMenuItem.Name = "extractFileToolStripMenuItem";
            this.extractFileToolStripMenuItem.Size = new Size(126, 22);
            this.extractFileToolStripMenuItem.Text = "Extract file";
            this.extractFileToolStripMenuItem.Click += extractFileToolStripMenuItem_Click;

            this.contextMenuStripMulti.Items.AddRange(new ToolStripItem[] { this.extractFilesToolStripMenuItem });
            this.contextMenuStripMulti.Name = "contextMenuStripMulti";
            this.contextMenuStripMulti.Size = new Size(132, 26);
            this.extractFilesToolStripMenuItem.Name = "extractFilesToolStripMenuItem";
            this.extractFilesToolStripMenuItem.Size = new Size(131, 22);
            this.extractFilesToolStripMenuItem.Text = "Extract files";
            this.extractFilesToolStripMenuItem.Click += extractFilesToolStripMenuItem_Click;

            this.saveFileDialog.Filter = "All files (*.*)|*.*";

            // Form
            this.AutoScaleDimensions = new SizeF(6f, 13f);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(729, 395);
            this.Controls.Add(this.toolStripContainer1);
            this.Controls.Add(this.statusStrip);
            this.Controls.Add(this.menuStrip);
            this.MainMenuStrip = this.menuStrip;
            this.Name = "MainForm";
            this.Text = "wxPirs";
            this.FormClosing += MainForm_FormClosing;
            this.Load += MainForm_Load;

            // Resume layouts
            this.menuStrip.ResumeLayout(false);
            this.menuStrip.PerformLayout();
            this.statusStrip.ResumeLayout(false);
            this.statusStrip.PerformLayout();
            this.toolStripContainer1.ContentPanel.ResumeLayout(false);
            this.toolStripContainer1.TopToolStripPanel.ResumeLayout(false);
            this.toolStripContainer1.TopToolStripPanel.PerformLayout();
            this.toolStripContainer1.ResumeLayout(false);
            this.toolStripContainer1.PerformLayout();
            this.splitContainerH.Panel1.ResumeLayout(false);
            this.splitContainerH.Panel2.ResumeLayout(false);
            this.splitContainerH.Panel2.PerformLayout();
            this.splitContainerH.ResumeLayout(false);
            this.splitContainerV.Panel1.ResumeLayout(false);
            this.splitContainerV.Panel2.ResumeLayout(false);
            this.splitContainerV.ResumeLayout(false);
            this.splitContainerList.Panel1.ResumeLayout(false);
            this.splitContainerList.Panel2.ResumeLayout(false);
            this.splitContainerList.ResumeLayout(false);
            this.fileToolStrip.ResumeLayout(false);
            this.fileToolStrip.PerformLayout();
            this.contextMenuStripFolder.ResumeLayout(false);
            this.contextMenuStrip.ResumeLayout(false);
            this.contextMenuStripMulti.ResumeLayout(false);
            this.contextMenuQueue.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        // set Out Folder button handler
        private void setOutFolderToolStripButton_Click(object sender, EventArgs e)
        {
            if (this.folderBrowserDialog.ShowDialog() != DialogResult.OK)
                return;
            this.outFolderPath = this.folderBrowserDialog.SelectedPath;
            this.toolStripStatusLabelOutFolder.Text = "Out Folder: " + this.outFolderPath;
        }

        // Batch (Process queue) extract all queued files into Out Folder
        private void batchToolStripButton_Click(object sender, EventArgs e)
        {
            if (this.listBoxQueue.Items.Count == 0)
            {
                this.log("Queue is empty\r\n");
                return;
            }

            // Prepare queue statuses if they are out of sync
            if (this.queueStatuses == null) this.queueStatuses = new List<QueueItemStatus>();
            while (this.queueStatuses.Count < this.listBoxQueue.Items.Count)
                this.queueStatuses.Add(QueueItemStatus.Pending);
            if (this.queueStatuses.Count > this.listBoxQueue.Items.Count)
                this.queueStatuses.RemoveRange(this.listBoxQueue.Items.Count, this.queueStatuses.Count - this.listBoxQueue.Items.Count);

            // Prepare batch logging - attempt to create log next to executable, fallback to LocalAppData
            string logFileName = $"batch_errors_{DateTime.Now:yyyyMMdd_HHmmss}.log";
            try
            {
                string exeDir = null;
                try
                {
                    // Assembly.Location can be empty for shadow-copied contexts; fall back to AppDomain base dir
                    exeDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                }
                catch
                {
                    exeDir = null;
                }
                if (string.IsNullOrEmpty(exeDir))
                    exeDir = AppDomain.CurrentDomain.BaseDirectory ?? Environment.CurrentDirectory;

                string candidate = Path.Combine(exeDir, logFileName);

                // Test write permission by opening with append (creates file if needed)
                try
                {
                    using (var fsTest = new FileStream(candidate, FileMode.Append, FileAccess.Write, FileShare.Read))
                    {
                        // write a small marker to indicate session start (will be appended later by StreamWriter)
                    }
                    batchLogFilePath = candidate;
                }
                catch
                {
                    // fallback to LocalAppData
                    string logDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "wxPirs");
                    try { Directory.CreateDirectory(logDir); } catch { }
                    batchLogFilePath = Path.Combine(logDir, logFileName);
                }
            }
            catch
            {
                // ultimate fallback
                string logDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "wxPirs");
                try { Directory.CreateDirectory(logDir); } catch { }
                batchLogFilePath = Path.Combine(logDir, logFileName);
            }

            // Inform user where errors will be written
            this.log($"Error log: {batchLogFilePath}\r\n");

            // Prepare UI
            this.toolStripProgressBar.Visible = true;
            this.toolStripProgressBar.Maximum = Math.Max(1, this.listBoxQueue.Items.Count);
            this.toolStripProgressBar.Value = 0;
            this.stopBatchToolStripButton.Enabled = true;
            this.batchToolStripButton.Enabled = false;
            this.stopBatchRequested = false;
            this.isBatchRunning = true;

            // Use UTF8 for log file
            using (StreamWriter sw = new StreamWriter(batchLogFilePath, true, Encoding.UTF8))
            {
                sw.WriteLine($"=== Batch started: {DateTime.Now:yyyy-MM-dd HH:mm:ss} ===");
                for (int i = 0; i < this.listBoxQueue.Items.Count; ++i)
                {
                    if (this.stopBatchRequested)
                    {
                        this.log("Batch stopped after current item.\r\n");
                        sw.WriteLine("Batch stopped by user request.");
                        break;
                    }

                    string file = this.listBoxQueue.Items[i].ToString();
                    this.log(string.Format("Processing: {0}\r\n", file));
                    sw.WriteLine($"Processing: {file}");
                    string savedLog = this.textBoxLog.Text;

                    // assume pending at start of processing
                    this.SetQueueStatus(i, QueueItemStatus.Pending);

                    try
                    {
                        this.openFile(file);
                    }
                    catch (Exception ex)
                    {
                        // Log to UI and file, mark item error, continue to next item
                        string err = string.Format("Error opening file {0}: {1}\r\n{2}\r\n", file, ex.Message, ex.ToString());
                        this.log(err);
                        try { sw.WriteLine(err); } catch { }
                        this.SetQueueStatus(i, QueueItemStatus.Error);
                        // update progress as done for this item
                        this.toolStripProgressBar.Value = Math.Min(this.toolStripProgressBar.Maximum, this.toolStripProgressBar.Value + 1);
                        continue;
                    }
                    // restore previous log (openFile clears it)
                    this.textBoxLog.Text = savedLog + this.textBoxLog.Text;

                    if (this.treeView.Nodes.Count == 0)
                    {
                        this.log("No content to extract\r\n");
                        try { sw.WriteLine("No content to extract"); } catch { }
                        this.SetQueueStatus(i, QueueItemStatus.Error);
                        // update progress as done
                        this.toolStripProgressBar.Value = Math.Min(this.toolStripProgressBar.Maximum, this.toolStripProgressBar.Value + 1);
                        continue;
                    }

                    // Use title from the opened file to name the subfolder if available, otherwise filename
                    string title = this.getTitle();
                    string rawName = string.IsNullOrWhiteSpace(title) ? Path.GetFileNameWithoutExtension(file) : title;

                    string targetBase = this.outFolderPath;
                    if (string.IsNullOrEmpty(targetBase))
                    {
                        // ask for out folder if not set
                        if (this.folderBrowserDialog.ShowDialog() != DialogResult.OK)
                        {
                            this.log("Out folder not set. Skipping.\r\n");
                            try { sw.WriteLine($"Out folder not set. Skipping {file}"); } catch { }
                            this.SetQueueStatus(i, QueueItemStatus.Error);
                            // update progress as done
                            this.toolStripProgressBar.Value = Math.Min(this.toolStripProgressBar.Maximum, this.toolStripProgressBar.Value + 1);
                            continue;
                        }
                        this.outFolderPath = this.folderBrowserDialog.SelectedPath;
                        this.toolStripStatusLabelOutFolder.Text = "Out Folder: " + this.outFolderPath;
                        targetBase = this.outFolderPath;
                    }

                    // sanitize and truncate taking the base path into account
                    string subdir = SanitizeFolderName(targetBase, rawName);

                    string targetDir = Path.Combine(targetBase, subdir);
                    try
                    {
                        // If the subfolder exists, delete it and recreate as requested
                        if (Directory.Exists(targetDir))
                        {
                            try
                            {
                                Directory.Delete(targetDir, true);
                            }
                            catch (Exception ex)
                            {
                                string err = string.Format("Error deleting existing output directory {0}: {1}\r\n{2}\r\n", targetDir, ex.Message, ex.ToString());
                                this.log(err);
                                try { sw.WriteLine(err); } catch { }
                                this.SetQueueStatus(i, QueueItemStatus.Error);
                                // Attempt to continue by skipping this item
                                // update progress as done
                                this.toolStripProgressBar.Value = Math.Min(this.toolStripProgressBar.Maximum, this.toolStripProgressBar.Value + 1);
                                continue;
                            }
                        }
                        Directory.CreateDirectory(targetDir);
                    }
                    catch (IOException ex)
                    {
                        string err = string.Format("Error creating output directory for {0}: {1}\r\n{2}\r\n", file, ex.Message, ex.ToString());
                        this.log(err);
                        try { sw.WriteLine(err); } catch { }
                        this.SetQueueStatus(i, QueueItemStatus.Error);
                        // update progress as done
                        this.toolStripProgressBar.Value = Math.Min(this.toolStripProgressBar.Maximum, this.toolStripProgressBar.Value + 1);
                        continue;
                    }

                    // extract starting from root parent (ushort.MaxValue)
                    try
                    {
                        this.extractFolder(ushort.MaxValue, subdir, targetBase);
                        this.log(string.Format("Done: {0}\r\n", file));
                        try { sw.WriteLine($"Done: {file}"); } catch { }
                        this.SetQueueStatus(i, QueueItemStatus.Success);
                    }
                    catch (Exception ex)
                    {
                        string err = string.Format("Error extracting {0}: {1}\r\n{2}\r\n", file, ex.Message, ex.ToString());
                        this.log(err);
                        try { sw.WriteLine(err); } catch { }
                        this.SetQueueStatus(i, QueueItemStatus.Error);
                    }
                    // increment progress bar after finishing this queue item
                    this.toolStripProgressBar.Value = Math.Min(this.toolStripProgressBar.Maximum, this.toolStripProgressBar.Value + 1);

                    Application.DoEvents();
                }

                sw.WriteLine($"=== Batch finished: {DateTime.Now:yyyy-MM-dd HH:mm:ss} ===");
            } // using StreamWriter

            // Reset UI
            this.isBatchRunning = false;
            this.stopBatchRequested = false;
            this.stopBatchToolStripButton.Enabled = false;
            this.batchToolStripButton.Enabled = true;
            this.toolStripProgressBar.Visible = false;
            this.log($"Batch finished. Error log: {batchLogFilePath}\r\n");
        }

        // Stop button handler: allow current item to finish but stop any following items
        private void stopBatchToolStripButton_Click(object sender, EventArgs e)
        {
            if (!this.isBatchRunning)
                return;
            this.stopBatchRequested = true;
            this.stopBatchToolStripButton.Enabled = false;
            this.log("Stop requested. Current item will finish, then batch will stop.\r\n");
        }

        // helper: set queue item status and refresh display
        private void SetQueueStatus(int index, QueueItemStatus status)
        {
            if (this.queueStatuses == null) this.queueStatuses = new List<QueueItemStatus>();
            if (index < 0) return;
            // ensure list long enough
            while (this.queueStatuses.Count <= index)
                this.queueStatuses.Add(QueueItemStatus.Pending);
            this.queueStatuses[index] = status;
            // refresh single item
            try
            {
                this.listBoxQueue.Invalidate(this.listBoxQueue.GetItemRectangle(index));
            }
            catch
            {
                this.listBoxQueue.Invalidate();
            }
        }

        // DrawItem handler to color items: Green for success, red for error, default otherwise
        private void listBoxQueue_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;
            e.Graphics.FillRectangle(SystemBrushes.Window, e.Bounds);

            bool selected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;

            // determine background color
            Color backColor = SystemColors.Window;
            Color foreColor = SystemColors.ControlText;

            if (!selected)
            {
                if (this.queueStatuses != null && e.Index < this.queueStatuses.Count)
                {
                    QueueItemStatus st = this.queueStatuses[e.Index];
                    if (st == QueueItemStatus.Success)
                        backColor = Color.Green; // changed to orange per request
                    else if (st == QueueItemStatus.Error)
                        backColor = Color.LightCoral;
                }
            }
            else
            {
                backColor = SystemColors.Highlight;
                foreColor = SystemColors.HighlightText;
            }

            using (SolidBrush b = new SolidBrush(backColor))
            {
                e.Graphics.FillRectangle(b, e.Bounds);
            }

            string text = this.listBoxQueue.Items[e.Index].ToString();
            Rectangle rect = e.Bounds;
            rect.Offset(2, 0);
            TextRenderer.DrawText(e.Graphics, text, e.Font, rect, foreColor, TextFormatFlags.VerticalCenter | TextFormatFlags.Left);

            e.DrawFocusRectangle();
        }

        // Queue context menu handling
        private void listBoxQueue_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                int idx = this.listBoxQueue.IndexFromPoint(e.Location);
                if (idx >= 0 && idx < this.listBoxQueue.Items.Count)
                {
                    this.listBoxQueue.SelectedIndex = idx;
                    this.contextMenuQueue.Show((Control)this.listBoxQueue, e.Location);
                }
            }
        }

        private void removeQueueToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int idx = this.listBoxQueue.SelectedIndex;
            if (idx >= 0)
            {
                this.listBoxQueue.Items.RemoveAt(idx);
                if (this.queueStatuses != null && idx < this.queueStatuses.Count)
                    this.queueStatuses.RemoveAt(idx);
            }
        }

        private void log(string message)
        {
            this.textBoxLog.AppendText(message);
        }

        private void openFile(object sender, EventArgs e)
        {
            if (this.openFileDialog.ShowDialog() != DialogResult.OK)
                return;
            this.openFile(this.openFileDialog.FileName);
        }

        private void openFile(string filename)
        {
            this.textBoxLog.Clear();
            this.folderBrowserDialog.SelectedPath = this.wr.extractFolderName(this.openFileDialog.FileName);
            if (this.br != null)
                this.br.Close();
            if (this.fs != null)
                this.fs.Dispose();
            this.treeView.BeginUpdate();
            this.treeView.Nodes.Clear();
            this.treeView.EndUpdate();
            this.listView.BeginUpdate();
            this.listView.Items.Clear();
            this.listView.EndUpdate();
            try
            {
                this.fs = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.Read);
                this.br = new BinaryReader((Stream)this.fs);
                this.getDescription();
                this.br.BaseStream.Seek(0L, SeekOrigin.Begin);
                int num1 = this.wr.readInt32(this.br);
                if (num1 != MainForm.MAGIC_PIRS && num1 != MainForm.MAGIC_LIVE && num1 != MainForm.MAGIC_CON_)
                {
                    this.log("Not a PIRS/LIVE file!\r\n");
                    this.br.Close();
                    this.fs.Close();
                }
                else
                {
                    this.br.BaseStream.Seek(49200L, SeekOrigin.Begin);
                    int num2 = this.wr.readInt32(this.br);
                    if (num1 == MainForm.MAGIC_CON_)
                    {
                        this.pirs_offset = MainForm.PIRS_TYPE2;
                        this.pirs_start = 49152L;
                    }
                    else if (num2 == (int)ushort.MaxValue)
                    {
                        this.pirs_offset = MainForm.PIRS_TYPE1;
                        this.pirs_start = MainForm.PIRS_BASE + this.pirs_offset;
                    }
                    else
                    {
                        this.pirs_offset = MainForm.PIRS_TYPE2;
                        this.pirs_start = MainForm.PIRS_BASE + this.pirs_offset;
                    }

                    //GetSongDta();
                    this.parse(filename);
                }
            }
            catch (IOException ex)
            {
                this.log(string.Format("{0}\r\n", (object)ex.Message));
            }
        }

        // New: open folder UI handler
        private void openFolderMenu_Click(object sender, EventArgs e)
        {
            if (this.folderBrowserDialog.ShowDialog() != DialogResult.OK)
                return;
            string folder = this.folderBrowserDialog.SelectedPath;
            try
            {
                string[] files = Directory.GetFiles(folder);
                this.listBoxQueue.BeginUpdate();
                this.listBoxQueue.Items.Clear();
                // reset statuses
                this.queueStatuses.Clear();
                foreach (string f in files)
                {
                    this.listBoxQueue.Items.Add(f);
                    this.queueStatuses.Add(QueueItemStatus.Pending);
                    this.log(string.Format("Queued: {0}\r\n", (object)f));
                }
                this.listBoxQueue.EndUpdate();
            }
            catch (IOException ex)
            {
                this.log(string.Format("Error reading folder: {0}\r\n", (object)ex.Message));
            }
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Exit();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Exit();
        }

        private void Exit()
        {
            base.Dispose();
        }

        private MainForm.PirsEntry getEntry()
        {
            MainForm.PirsEntry pirsEntry = new MainForm.PirsEntry();
            pirsEntry.Filename = this.wr.readString(this.br, 38U);
            if (pirsEntry.Filename.Trim() == "")
                return pirsEntry;
            pirsEntry.Unknow = this.wr.readInt32(this.br);
            pirsEntry.BlockLen = this.wr.readInt32(this.br);
            pirsEntry.Cluster = this.br.ReadInt32() >> 8;
            pirsEntry.Parent = this.wr.readUInt16(this.br);
            pirsEntry.Size = this.wr.readInt32(this.br);
            pirsEntry.DateTime1 = this.dosDateTime(this.wr.readInt32(this.br));
            pirsEntry.DateTime2 = this.dosDateTime(this.wr.readInt32(this.br));
            return pirsEntry;
        }

        private DateTime dosDateTime(int datetime)
        {
            return DateTime.Now;
        }

        private DateTime dosDateTime(short date, short time)
        {
            return DateTime.Now;
        }

        private void getFiles(TreeNode tn)
        {
            int num1 = 0;
            while (true)
            {
                this.br.BaseStream.Seek(this.pirs_start + (long)(num1 * 64), SeekOrigin.Begin);
                MainForm.PirsEntry entry = this.getEntry();
                if (!(entry.Filename.Trim() == ""))
                {
                    if (entry.Cluster == 0 && entry.Size == 0 && (int)entry.Parent == (int)Convert.ToUInt16(tn.Tag))
                    {
                        ListViewItem listViewItem = new ListViewItem();
                        listViewItem.Text = entry.Filename;
                        listViewItem.SubItems.Add(entry.Size.ToString());
                        listViewItem.SubItems.Add(entry.Cluster.ToString());
                        listViewItem.SubItems.Add(entry.DateTime1.ToString());
                        listViewItem.SubItems.Add("");
                        listViewItem.ImageIndex = 0;
                        listViewItem.Tag = (object)num1;
                        this.listView.Items.Add(listViewItem);
                    }
                    ++num1;
                }
                else
                    break;
            }
            int num2 = 0;
            while (true)
            {
                this.br.BaseStream.Seek(this.pirs_start + (long)(num2 * 64), SeekOrigin.Begin);
                MainForm.PirsEntry pirsEntry = this.getEntry();
                if (!(pirsEntry.Filename.Trim() == ""))
                {
                    if (pirsEntry.Cluster != 0 && (int)pirsEntry.Parent == (int)Convert.ToUInt16(tn.Tag))
                    {
                        ListViewItem listViewItem = new ListViewItem();
                        listViewItem.Text = pirsEntry.Filename;
                        listViewItem.SubItems.Add(pirsEntry.Size.ToString());
                        listViewItem.SubItems.Add(pirsEntry.Cluster.ToString());
                        listViewItem.SubItems.Add(pirsEntry.DateTime1.ToString());
                        listViewItem.SubItems.Add("");
                        listViewItem.ImageIndex = 1;
                        listViewItem.ToolTipText = string.Format("Offset :0x{0:X8}", (object)this.getOffset((long)pirsEntry.Cluster));
                        this.listView.Items.Add(listViewItem);
                    }
                    ++num2;
                }
                else
                    break;
            }
        }

        private void getDirectories(TreeNode tn)
        {
            int num = 0;
            while (true)
            {
                this.br.BaseStream.Seek(this.pirs_start + (long)(num * 64), SeekOrigin.Begin);
                MainForm.PirsEntry entry = this.getEntry();
                if (!(entry.Filename.Trim() == ""))
                {
                    if (entry.Size == 0 && entry.Cluster == 0 && (int)entry.Parent == (int)Convert.ToUInt16(tn.Tag))
                    {
                        TreeNode treeNode = new TreeNode(entry.Filename);
                        treeNode.Tag = (object)num;
                        treeNode.ToolTipText = string.Format("0x{0:X4}", treeNode.Tag);
                        tn.Nodes.Add(treeNode);
                        this.getDirectories(treeNode);
                    }
                    ++num;
                }
                else
                    break;
            }
        }

        private void parse(string filename)
        {
            this.treeView.BeginUpdate();
            this.treeView.Nodes.Clear();
            this.listView.BeginUpdate();
            this.listView.Items.Clear();
            TreeNode treeNode = new TreeNode(this.wr.extractFileName(filename), 0, 0);
            treeNode.Tag = (object)ushort.MaxValue;
            treeNode.ToolTipText = string.Format("0x{0:X4}", treeNode.Tag);
            this.treeView.Nodes.Add(treeNode);
            this.getDirectories(treeNode);
            this.getFiles(treeNode);
            treeNode.Expand();
            this.listView.EndUpdate();
            this.treeView.EndUpdate();
        }

        private void treeView_AfterSelect(object sender, TreeViewEventArgs e)
        {
            this.listView.BeginUpdate();
            this.listView.Items.Clear();
            this.getFiles(e.Node);
            this.listView.EndUpdate();
        }

        private long getOffset(long cluster)
        {
            long num1 = this.pirs_start + cluster * 4096L;
            long num2 = cluster / 170L;
            long num3 = num2 / 170L;
            if (num2 > 0L)
                num1 += (num2 + 1L) * this.pirs_offset;
            if (num3 > 0L)
                num1 += (num3 + 1L) * this.pirs_offset;
            return num1;
        }

        private void extractFile(ListViewItem listViewItem, string filename)
        {
            try
            {
                if (!Directory.Exists(this.wr.extractFolderName(filename)))
                    Directory.CreateDirectory(this.wr.extractFolderName(filename));
            }
            catch (IOException ex)
            {
                this.log(string.Format("Error creating directory: {0}\r\n", ex.Message));
            }
            FileStream fileStream;
            BinaryWriter binaryWriter;
            try
            {
                fileStream = new FileStream(filename, FileMode.Create, FileAccess.Write, FileShare.None);
                binaryWriter = new BinaryWriter((Stream)fileStream);
            }
            catch (IOException ex)
            {
                this.log(string.Format("Error : {0}\r\n", (object)ex));
                return;
            }
            long cluster1 = Convert.ToInt64(listViewItem.SubItems[this.columnHeaderCluster.Index].Text);
            this.getOffset(cluster1);
            long num1 = Convert.ToInt64(listViewItem.SubItems[this.columnHeaderSize.Index].Text);
            long num2 = num1 >> 12;
            long num3 = num1 - (num2 << 12);
            for (long cluster2 = cluster1; cluster2 < cluster1 + num2; ++cluster2)
            {
                this.br.BaseStream.Seek(this.getOffset(cluster2), SeekOrigin.Begin);
                binaryWriter.Write(this.br.ReadBytes(4096));
            }
            this.br.BaseStream.Seek(this.getOffset(cluster1 + num2), SeekOrigin.Begin);
            binaryWriter.Write(this.br.ReadBytes((int)num3));
            listViewItem.SubItems[this.columnHeaderStatus.Index].Text = "Done";
            Application.DoEvents();
            binaryWriter.Close();
            fileStream.Dispose();
        }

        private void extractFile(long cluster, long size, string filename)
        {
            try
            {
                if (!Directory.Exists(this.wr.extractFolderName(filename)))
                    Directory.CreateDirectory(this.wr.extractFolderName(filename));
            }
            catch (IOException ex)
            {
                this.log(string.Format("Error creating directory: {0}\r\n", ex.Message));
            }
            FileStream fileStream;
            BinaryWriter binaryWriter;
            try
            {
                fileStream = new FileStream(filename, FileMode.Create, FileAccess.Write, FileShare.None);
                binaryWriter = new BinaryWriter((Stream)fileStream);
            }
            catch (IOException ex)
            {
                this.log(string.Format("Error : {0}\r\n", (object)ex));
                return;
            }
            long num1 = size >> 12;
            long num2 = size - (num1 << 12);
            for (long cluster1 = cluster; cluster1 < cluster + num1; ++cluster1)
            {
                this.br.BaseStream.Seek(this.getOffset(cluster1), SeekOrigin.Begin);
                binaryWriter.Write(this.br.ReadBytes(4096));
                Application.DoEvents();
            }
            this.br.BaseStream.Seek(this.getOffset(cluster + num1), SeekOrigin.Begin);
            binaryWriter.Write(this.br.ReadBytes((int)num2));
            Application.DoEvents();
            binaryWriter.Close();
            fileStream.Dispose();
        }

        private void extractFolder(ListViewItem listViewItem, string pathname)
        {
            listViewItem.SubItems[this.columnHeaderStatus.Index].Text = "Working...";
            Application.DoEvents();
            this.extractFolder(Convert.ToUInt16(listViewItem.Tag), listViewItem.Text, pathname);
            listViewItem.SubItems[this.columnHeaderStatus.Index].Text = "Done";
            Application.DoEvents();
        }

        private void extractFolder(ushort tag, string foldername, string pathname)
        {
            ushort tag1 = (ushort)0;
            while (true)
            {
                this.br.BaseStream.Seek(MainForm.PIRS_BASE + this.pirs_offset + (long)((int)tag1 * 64), SeekOrigin.Begin);
                MainForm.PirsEntry pirsEntry = new MainForm.PirsEntry();
                pirsEntry = this.getEntry();
                if (!(pirsEntry.Filename.Trim() == ""))
                {
                    if (pirsEntry.Cluster == 0 && pirsEntry.Size == 0 && (int)pirsEntry.Parent == (int)tag)
                        this.extractFolder(tag1, pirsEntry.Filename, pathname + "\\" + foldername);
                    else if (pirsEntry.Cluster != 0 && (int)pirsEntry.Parent == (int)tag)
                        this.extractFile((long)pirsEntry.Cluster, (long)pirsEntry.Size, pathname + "\\" + foldername + "\\" + pirsEntry.Filename);
                    ++tag1;
                }
                else
                    break;
            }
        }

        private void listView_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right && this.listView.SelectedItems.Count == 1)
            {
                if (this.isFolder(this.listView.SelectedItems[0]))
                    this.contextMenuStripMulti.Show((Control)this.listView, e.X, e.Y);
                else
                    this.contextMenuStrip.Show((Control)this.listView, e.X, e.Y);
            }
            else
            {
                if (e.Button != MouseButtons.Right || this.listView.SelectedItems.Count <= 1)
                    return;
                this.contextMenuStripMulti.Show((Control)this.listView, e.X, e.Y);
            }
        }

        private bool isFolder(ListViewItem listViewItem)
        {
            return listViewItem.ImageIndex == 0;
        }

        private void extractFilesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.folderBrowserDialog.ShowDialog() != DialogResult.OK)
                return;
            for (int index = 0; index < this.listView.SelectedItems.Count; ++index)
            {
                if (this.isFolder(this.listView.SelectedItems[index]))
                {
                    this.log(string.Format("Extract folder {0}\r\n", (object)this.listView.SelectedItems[index].Text));
                    this.extractFolder(this.listView.SelectedItems[index], this.folderBrowserDialog.SelectedPath);
                }
                else
                    this.extractFile(this.listView.SelectedItems[index], this.folderBrowserDialog.SelectedPath + "\\" + this.listView.SelectedItems[index].Text);
                Application.DoEvents();
            }
        }

        private void extractFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.saveFileDialog.FileName = this.listView.SelectedItems[0].Text;
            if (this.saveFileDialog.ShowDialog() != DialogResult.OK)
                return;
            this.extractFile(this.listView.SelectedItems[0], this.saveFileDialog.FileName);
        }

        private void aboutPirsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutBox aboutBox = new AboutBox();
            int num = (int)aboutBox.ShowDialog();
            aboutBox.Dispose();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            this.toolStripStatusLabelVersion.Text = Application.ProductName + " - " + ((object)Assembly.GetExecutingAssembly().GetName().Version).ToString();
            if (this.args.Length <= 0)
                return;
            this.openFileDialog.FileName = this.args[0];
            this.openFile(this.args[0]);
        }

        private void treeView_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Button != MouseButtons.Right)
                return;
            this.treeView.SelectedNode = e.Node;
            this.contextMenuStripFolder.Show((Control)this.treeView, e.X, e.Y);
        }

        private void extractFolder(object sender, EventArgs e)
        {
            if (this.treeView.SelectedNode == null || this.folderBrowserDialog.ShowDialog() != DialogResult.OK)
                return;
            for (int index = 0; index < this.listView.Items.Count; ++index)
            {
                if (this.isFolder(this.listView.Items[index]))
                    this.extractFolder(this.listView.Items[index], this.folderBrowserDialog.SelectedPath);
                else
                    this.extractFile(this.listView.Items[index], this.folderBrowserDialog.SelectedPath + "\\" + this.listView.Items[index].Text);
                Application.DoEvents();
            }
        }

        private void extractAll(object sender, EventArgs e)
        {
            this.treeView.SelectedNode = this.treeView.Nodes[0];
            this.extractFolder(sender, e);
        }

        private long getCultureOffset()
        {
            return !(Application.CurrentCulture.TwoLetterISOLanguageName.ToLower() == "de") ? (!(Application.CurrentCulture.TwoLetterISOLanguageName.ToLower() == "fr") ? (!(Application.CurrentCulture.TwoLetterISOLanguageName.ToLower() == "es") ? (!(Application.CurrentCulture.TwoLetterISOLanguageName.ToLower() == "it") ? 0L : 1280L) : 1024L) : 768L) : 512L;
        }

        private void getDescription()
        {
            long cultureOffset = this.getCultureOffset();
            this.br.BaseStream.Seek(1040L + cultureOffset, SeekOrigin.Begin);
            byte[] numArray1 = new byte[256];
            this.log("Title : " + this.wr.unicodeToStr(this.br.ReadBytes(256), 2) + "\r\n");
            this.br.BaseStream.Seek(3344L + cultureOffset, SeekOrigin.Begin);
            byte[] numArray2 = new byte[256];
            this.log("Description : " + this.wr.unicodeToStr(this.br.ReadBytes(256), 2) + "\r\n");
            this.br.BaseStream.Seek(5648L, SeekOrigin.Begin);
            byte[] numArray3 = new byte[256];
            this.log("Publisher : " + this.wr.unicodeToStr(this.br.ReadBytes(256), 2) + "\r\n");
        }

        // New: double click queued file to open
        private void listBoxQueue_DoubleClick(object sender, EventArgs e)
        {
            if (this.listBoxQueue.SelectedItem == null)
                return;
            string file = this.listBoxQueue.SelectedItem.ToString();
            if (File.Exists(file))
            {
                this.openFile(file);
            }
            else
            {
                this.log(string.Format("File not found: {0}\r\n", (object)file));
            }
        }

        // New: helper method to get the Title from the opened file
        private string getTitle()
        {
            // If no file/reader opened, return empty
            if (this.br == null || this.br.BaseStream == null)
                return string.Empty;

            long originalPos = -1;
            try
            {
                // Save current stream position (best-effort)
                try
                {
                    originalPos = this.br.BaseStream.Position;
                }
                catch
                {
                    originalPos = -1;
                }

                // Seek to the same offset used in getDescription()
                long cultureOffset = this.getCultureOffset();
                long titleOffset = 1040L + cultureOffset;
                this.br.BaseStream.Seek(titleOffset, SeekOrigin.Begin);

                // Read 256 bytes and convert using existing reader helper
                byte[] buffer = this.br.ReadBytes(256);
                string title = this.wr.unicodeToStr(buffer, 2);

                // Trim and return safely
                return (title ?? string.Empty).Trim();
            }
            catch
            {
                // On any error return empty string
                return string.Empty;
            }
            finally
            {
                // Restore original position if possible
                if (originalPos >= 0)
                {
                    try
                    {
                        this.br.BaseStream.Seek(originalPos, SeekOrigin.Begin);
                    }
                    catch
                    {
                        // ignore restoration errors
                    }
                }
            }
        }

        /* Replace the old SanitizeFolderName(string) with this more robust implementation.
           It preserves Unicode characters, replaces invalid filename characters, trims trailing dots/spaces,
           protects reserved Windows names, and truncates based on the base path so the full path won't exceed MAX_PATH.
        */
        private string SanitizeFolderName(string basePath, string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return string.Empty;

            // Normalize to composed form to help with file APIs and visually-sortable names
            name = name.Normalize(System.Text.NormalizationForm.FormC).Trim();

            // Replace invalid filename chars with underscores but keep non-ASCII characters (Japanese/Chinese/etc.)
            char[] invalid = Path.GetInvalidFileNameChars();
            foreach (char c in invalid)
            {
                name = name.Replace(c, '_');
            }

            // Remove trailing spaces and dots (Windows forbids those)
            name = name.TrimEnd(' ', '.');

            // Protect reserved device names (CON, PRN, AUX, NUL, COM1..COM9, LPT1..LPT9)
            var reserved = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "CON","PRN","AUX","NUL",
                "COM1","COM2","COM3","COM4","COM5","COM6","COM7","COM8","COM9",
                "LPT1","LPT2","LPT3","LPT4","LPT5","LPT6","LPT7","LPT8","LPT9"
            };
            if (reserved.Contains(name))
                name = "_" + name;

            // Now ensure the full path length will not exceed MAX_PATH (260) unless the app/OS supports long paths.
            const int MAX_PATH = 260;

            try
            {
                // Resolve base full path
                string baseFull = string.IsNullOrEmpty(basePath) ? Directory.GetCurrentDirectory() : Path.GetFullPath(basePath);

                // If baseFull is already long or invalid for GetFullPath, fall back to a conservative limit
                if (string.IsNullOrEmpty(baseFull))
                    baseFull = Directory.GetCurrentDirectory();

                // We need to leave room for a directory separator and possibly more path elements.
                int allowedForName = Math.Max(1, MAX_PATH - baseFull.Length - 1);

                if (name.Length > allowedForName)
                {
                    // Truncate preserving Unicode character boundaries (C# substring is by char so OK)
                    name = name.Substring(0, allowedForName);
                    // Trim any trailing spaces/dots introduced by truncation
                    name = name.TrimEnd(' ', '.');
                    if (string.IsNullOrEmpty(name))
                        name = Path.GetFileNameWithoutExtension(baseFull); // conservative fallback
                }
            }
            catch
            {
                // In case of any path resolution error, fallback to a safer fixed-length truncation
                if (name.Length > 200)
                    name = name.Substring(0, 200);
            }

            return name;
        }

        public struct PirsEntry
        {
            public string Filename;
            public int Unknow;
            public int BlockLen;
            public int Cluster;
            public ushort Parent;
            public int Size;
            public DateTime DateTime1;
            public DateTime DateTime2;
        }
    }
}
