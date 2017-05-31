/******************************************************
                   Simple MAPI.NET
		      netmaster@swissonline.ch
*******************************************************/

using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Win32Mapi;

namespace SimpleMAPIdotNET
{
	/// <summary>
	/// MainForm is the main window.
	/// </summary>
	public class MainForm : System.Windows.Forms.Form
	{
		private System.Windows.Forms.ListView listInbox;
		private System.Windows.Forms.ColumnHeader colinbxFrom;
		private System.Windows.Forms.ColumnHeader colinbxSubj;
		private System.Windows.Forms.ColumnHeader colinbxRecvd;
		private System.Windows.Forms.TextBox textMail;
		private System.Windows.Forms.Button buttonRefresh;
		private System.Windows.Forms.Button buttonDeleteMail;
		private System.Windows.Forms.Button buttonSendNew;
		private System.Windows.Forms.ImageList imageLstMail;
		private System.Windows.Forms.ComboBox comboAttachm;
		private System.Windows.Forms.Button buttonSaveAtt;
		private System.Windows.Forms.Splitter splitter1;
		private System.Windows.Forms.Panel panelBottom;
		private System.ComponentModel.IContainer components;

		private Mapi ma = new Mapi();
		private bool first_activated = false;
		private Font boldFont;
		MailEnvelop currentMail;
		MailComparer comparer = new MailComparer();

		public MainForm()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			boldFont = new Font( listInbox.Font, listInbox.Font.Style | FontStyle.Bold );
			listInbox.ListViewItemSorter = comparer;
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(MainForm));
			this.textMail = new System.Windows.Forms.TextBox();
			this.buttonSendNew = new System.Windows.Forms.Button();
			this.listInbox = new System.Windows.Forms.ListView();
			this.colinbxFrom = new System.Windows.Forms.ColumnHeader();
			this.colinbxSubj = new System.Windows.Forms.ColumnHeader();
			this.colinbxRecvd = new System.Windows.Forms.ColumnHeader();
			this.imageLstMail = new System.Windows.Forms.ImageList(this.components);
			this.buttonRefresh = new System.Windows.Forms.Button();
			this.buttonDeleteMail = new System.Windows.Forms.Button();
			this.comboAttachm = new System.Windows.Forms.ComboBox();
			this.buttonSaveAtt = new System.Windows.Forms.Button();
			this.panelBottom = new System.Windows.Forms.Panel();
			this.splitter1 = new System.Windows.Forms.Splitter();
			this.panelBottom.SuspendLayout();
			this.SuspendLayout();
			// 
			// textMail
			// 
			this.textMail.Dock = System.Windows.Forms.DockStyle.Fill;
			this.textMail.HideSelection = false;
			this.textMail.Location = new System.Drawing.Point(0, 134);
			this.textMail.MaxLength = 256000;
			this.textMail.Multiline = true;
			this.textMail.Name = "textMail";
			this.textMail.ScrollBars = System.Windows.Forms.ScrollBars.Both;
			this.textMail.Size = new System.Drawing.Size(600, 231);
			this.textMail.TabIndex = 3;
			this.textMail.Text = "-";
			this.textMail.WordWrap = false;
			// 
			// buttonSendNew
			// 
			this.buttonSendNew.Anchor = (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right);
			this.buttonSendNew.Location = new System.Drawing.Point(509, 4);
			this.buttonSendNew.Name = "buttonSendNew";
			this.buttonSendNew.Size = new System.Drawing.Size(80, 32);
			this.buttonSendNew.TabIndex = 4;
			this.buttonSendNew.Text = "Send Mail...";
			this.buttonSendNew.Click += new System.EventHandler(this.buttonSendNew_Click);
			// 
			// listInbox
			// 
			this.listInbox.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
																						this.colinbxFrom,
																						this.colinbxSubj,
																						this.colinbxRecvd});
			this.listInbox.Dock = System.Windows.Forms.DockStyle.Top;
			this.listInbox.FullRowSelect = true;
			this.listInbox.GridLines = true;
			this.listInbox.HideSelection = false;
			this.listInbox.MultiSelect = false;
			this.listInbox.Name = "listInbox";
			this.listInbox.Size = new System.Drawing.Size(600, 128);
			this.listInbox.SmallImageList = this.imageLstMail;
			this.listInbox.Sorting = System.Windows.Forms.SortOrder.Descending;
			this.listInbox.TabIndex = 2;
			this.listInbox.View = System.Windows.Forms.View.Details;
			this.listInbox.ColumnClick += new System.Windows.Forms.ColumnClickEventHandler(this.listInbox_ColumnClick);
			this.listInbox.SelectedIndexChanged += new System.EventHandler(this.listInbox_SelectedIndexChanged);
			// 
			// colinbxFrom
			// 
			this.colinbxFrom.Text = "From";
			this.colinbxFrom.Width = 160;
			// 
			// colinbxSubj
			// 
			this.colinbxSubj.Text = "Subject";
			this.colinbxSubj.Width = 296;
			// 
			// colinbxRecvd
			// 
			this.colinbxRecvd.Text = "Received";
			this.colinbxRecvd.Width = 140;
			// 
			// imageLstMail
			// 
			this.imageLstMail.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit;
			this.imageLstMail.ImageSize = new System.Drawing.Size(16, 16);
			this.imageLstMail.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageLstMail.ImageStream")));
			this.imageLstMail.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// buttonRefresh
			// 
			this.buttonRefresh.Anchor = (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left);
			this.buttonRefresh.Location = new System.Drawing.Point(6, 5);
			this.buttonRefresh.Name = "buttonRefresh";
			this.buttonRefresh.Size = new System.Drawing.Size(64, 32);
			this.buttonRefresh.TabIndex = 1;
			this.buttonRefresh.Text = "Refresh";
			this.buttonRefresh.Click += new System.EventHandler(this.buttonRefresh_Click);
			// 
			// buttonDeleteMail
			// 
			this.buttonDeleteMail.Anchor = (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left);
			this.buttonDeleteMail.Location = new System.Drawing.Point(82, 5);
			this.buttonDeleteMail.Name = "buttonDeleteMail";
			this.buttonDeleteMail.Size = new System.Drawing.Size(64, 32);
			this.buttonDeleteMail.TabIndex = 5;
			this.buttonDeleteMail.Text = "Delete";
			this.buttonDeleteMail.Click += new System.EventHandler(this.buttonDeleteMail_Click);
			// 
			// comboAttachm
			// 
			this.comboAttachm.Anchor = (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left);
			this.comboAttachm.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.comboAttachm.DropDownWidth = 320;
			this.comboAttachm.Enabled = false;
			this.comboAttachm.Location = new System.Drawing.Point(182, 9);
			this.comboAttachm.Name = "comboAttachm";
			this.comboAttachm.Size = new System.Drawing.Size(160, 21);
			this.comboAttachm.TabIndex = 6;
			// 
			// buttonSaveAtt
			// 
			this.buttonSaveAtt.Anchor = (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left);
			this.buttonSaveAtt.Enabled = false;
			this.buttonSaveAtt.Location = new System.Drawing.Point(343, 4);
			this.buttonSaveAtt.Name = "buttonSaveAtt";
			this.buttonSaveAtt.Size = new System.Drawing.Size(57, 32);
			this.buttonSaveAtt.TabIndex = 7;
			this.buttonSaveAtt.Text = "Save...";
			this.buttonSaveAtt.Click += new System.EventHandler(this.buttonSaveAtt_Click);
			// 
			// panelBottom
			// 
			this.panelBottom.Controls.AddRange(new System.Windows.Forms.Control[] {
																					  this.buttonRefresh,
																					  this.buttonDeleteMail,
																					  this.comboAttachm,
																					  this.buttonSaveAtt,
																					  this.buttonSendNew});
			this.panelBottom.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.panelBottom.Location = new System.Drawing.Point(0, 365);
			this.panelBottom.Name = "panelBottom";
			this.panelBottom.Size = new System.Drawing.Size(600, 40);
			this.panelBottom.TabIndex = 8;
			// 
			// splitter1
			// 
			this.splitter1.Dock = System.Windows.Forms.DockStyle.Top;
			this.splitter1.Location = new System.Drawing.Point(0, 128);
			this.splitter1.Name = "splitter1";
			this.splitter1.Size = new System.Drawing.Size(600, 6);
			this.splitter1.TabIndex = 9;
			this.splitter1.TabStop = false;
			// 
			// MainForm
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(600, 405);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.textMail,
																		  this.panelBottom,
																		  this.splitter1,
																		  this.listInbox});
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.MinimumSize = new System.Drawing.Size(512, 400);
			this.Name = "MainForm";
			this.Text = "Simple MAPI.NET";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.MainForm_Closing);
			this.Activated += new System.EventHandler(this.MainForm_Activated);
			this.panelBottom.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new MainForm());
		}






		/// <summary> Fill or refresh the Inbox </summary>
		private void RefreshInbox()
			{
			comparer.Column = 2;
			comparer.Order = SortOrder.Descending;
			currentMail = null;
			comboAttachm.Items.Clear();
			comboAttachm.Enabled = false;
			buttonSaveAtt.Enabled = false;
			textMail.Text = null;
			listInbox.BeginUpdate();
			listInbox.Items.Clear();
			
			ma.Reset();
			bool more = false;
			string[]	itemstrings = new string[ 3 ];
			do
				{
				MailEnvelop env = new MailEnvelop();
				more = ma.Next( ref env );
				if( more )
					{
					int imgi = 0;
					if( ! env.unread )
						imgi = 2;
					if( env.atts != 0 )
						imgi++;
						
					itemstrings[0] = env.from;
					itemstrings[1] = env.subject;
					itemstrings[2] = env.date.ToString();
					ListViewItem ii = new ListViewItem( itemstrings, imgi );
					ii.Tag = env;
					if( env.unread )
						ii.Font = boldFont;
					listInbox.Items.Add( ii );
					}
				}
			while( more );
			
			listInbox.EndUpdate();
			}

		/// <summary> Once at startup, do logon mapi and fill Inbox </summary>
		private void MainForm_Activated( object sender, System.EventArgs e )
			{
			if( ! first_activated )
				{
				first_activated = true;
				if( ma.Logon( this.Handle ) )
					RefreshInbox();
				}
			}

		/// <summary> User changed selection in Inbox. Show the mail content </summary>
		private void listInbox_SelectedIndexChanged( object sender, System.EventArgs e )
			{
			currentMail = null;
			comboAttachm.Items.Clear();
			comboAttachm.Enabled = false;
			buttonSaveAtt.Enabled = false;
			if( listInbox.SelectedItems.Count != 1 )		// no selection
				return;
			ListViewItem	selitem = listInbox.SelectedItems[0];
			currentMail = selitem.Tag as MailEnvelop;
			MailAttach[] aat;
			textMail.Text = ma.Read( currentMail.id, out aat );
			if( aat != null )
				{								// has attachment
				comboAttachm.BeginUpdate();				// update attachment list
				foreach( MailAttach a in aat )
					{
					if( a.name != null )
						comboAttachm.Items.Add( a.name );
					}
				comboAttachm.EndUpdate();
				if( comboAttachm.Items.Count > 0 )
					{
					comboAttachm.SelectedIndex = 0;
					comboAttachm.Enabled = true;
					buttonSaveAtt.Enabled = true;
					}
				}
			}

		private void buttonRefresh_Click(object sender, System.EventArgs e)
			{
			RefreshInbox();
			}

		private void MainForm_Closing(object sender, System.ComponentModel.CancelEventArgs e)
			{
			ma.Logoff();
			}


		/// <summary> User clicked button to send new mail </summary>
		private void buttonSendNew_Click(object sender, System.EventArgs e)
			{
			SendForm	frmSend = new SendForm( ref ma );
			frmSend.ShowDialog( this );
			}


		/// <summary> User clicked button to delete selected mail item </summary>
		private void buttonDeleteMail_Click(object sender, System.EventArgs e)
			{
			if( currentMail == null )		// no selection
				return;

			DialogResult r = MessageBox.Show( this, "are you sure?", "Delete Mail", MessageBoxButtons.YesNo, MessageBoxIcon.Question );
			if( r == DialogResult.Yes )
				{
				if( ! ma.Delete( currentMail.id ) )
					MessageBox.Show( this, "MAPIDeleteMail failed! " + ma.Error(), "Delete Mail", MessageBoxButtons.OK, MessageBoxIcon.Warning );
				RefreshInbox();
				}
			}

		/// <summary> User clicked button to save selected attachment </summary>
		private void buttonSaveAtt_Click(object sender, System.EventArgs e)
			{
			SaveFileDialog sd = new SaveFileDialog();
			sd.FileName = comboAttachm.Text;
			sd.Title = "Save attachment as...";
			sd.Filter = "All files (*.*)|*.*";
			sd.AddExtension = false;
			sd.FilterIndex = 1;
			sd.RestoreDirectory = true;
			if( sd.ShowDialog() != DialogResult.OK )
				return;

			if( ! ma.SaveAttachm( currentMail.id, comboAttachm.Text, sd.FileName ) )
				MessageBox.Show( this, "save failed! " + ma.Error(), "Save attachment", MessageBoxButtons.OK, MessageBoxIcon.Warning );
			}

		/// <summary> User clicked Inbox column, do sort items </summary>
		private void listInbox_ColumnClick(object sender, System.Windows.Forms.ColumnClickEventArgs e)
			{
			if( e.Column == comparer.Column )
				{
				if( comparer.Order == SortOrder.Ascending )
					comparer.Order = SortOrder.Descending;
				else
					comparer.Order = SortOrder.Ascending;
				}
			else if( (e.Column >= 0) & (e.Column <= 2) )
				comparer.Column = e.Column;
				
			listInbox.Sort();
			}


	}



/// <summary>
/// Class providing IComparer to sort email items in Inbox.
/// </summary>
internal class MailComparer : IComparer
	{
	public int Compare( object object1 , object object2 )
		{
		ListViewItem lv1 = object1 as ListViewItem;
		ListViewItem lv2 = object2 as ListViewItem;
		if( (lv1 == null) || (lv2 == null) )
			return 0; 

		MailEnvelop ev1 = lv1.Tag as MailEnvelop;
		MailEnvelop ev2 = lv2.Tag as MailEnvelop;
		if( (ev1 == null) || (ev2 == null) )
			return 0; 

		int r = 0;
		if( sortcolumn == 0 )
			r = String.Compare( ev1.from, ev2.from );
		else if( sortcolumn == 1 )
			r = String.Compare( ev1.subject, ev2.subject );
		else 
			r = DateTime.Compare( ev1.date, ev2.date );
		if( sorting == SortOrder.Descending )
			r = -r;
		return r;
		}

	public int Column
		{
		set { sortcolumn = value ; }
		get { return sortcolumn; }
		}

	public SortOrder Order
		{
		set { sorting = value; }
		get { return sorting; }
		}

	private int			sortcolumn = 2;
	private SortOrder	sorting = SortOrder.Descending;
	}

}
