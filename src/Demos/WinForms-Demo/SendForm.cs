/******************************************************
                   Simple MAPI.NET
		      netmaster@swissonline.ch
*******************************************************/

using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

using Win32Mapi;

namespace SimpleMAPIdotNET
{
	/// <summary>
	/// Summary description for SendForm.
	/// </summary>
	public class SendForm : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button buttonAddrTO;
		private System.Windows.Forms.TextBox textTO;
		private System.Windows.Forms.Button buttonAddrCC;
		private System.Windows.Forms.TextBox textCC;
		private System.Windows.Forms.Label labelSubj;
		private System.Windows.Forms.TextBox textSubject;
		private System.Windows.Forms.TextBox textMail;
		private System.Windows.Forms.Button buttonAttach;
		private System.Windows.Forms.ComboBox comboAttachm;
		private System.Windows.Forms.Button buttonSend;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		private Mapi ma;

		public SendForm( ref Mapi rma )
		{
			ma = rma;
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			ma.Reset();
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(SendForm));
			this.textTO = new System.Windows.Forms.TextBox();
			this.textSubject = new System.Windows.Forms.TextBox();
			this.labelSubj = new System.Windows.Forms.Label();
			this.textMail = new System.Windows.Forms.TextBox();
			this.textCC = new System.Windows.Forms.TextBox();
			this.buttonSend = new System.Windows.Forms.Button();
			this.comboAttachm = new System.Windows.Forms.ComboBox();
			this.buttonAttach = new System.Windows.Forms.Button();
			this.buttonAddrTO = new System.Windows.Forms.Button();
			this.buttonAddrCC = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// textTO
			// 
			this.textTO.Anchor = ((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right);
			this.textTO.Location = new System.Drawing.Point(64, 8);
			this.textTO.MaxLength = 5000;
			this.textTO.Name = "textTO";
			this.textTO.Size = new System.Drawing.Size(336, 20);
			this.textTO.TabIndex = 2;
			this.textTO.Text = "netmaster@swissonline.ch";
			// 
			// textSubject
			// 
			this.textSubject.Anchor = ((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right);
			this.textSubject.Location = new System.Drawing.Point(64, 104);
			this.textSubject.MaxLength = 300;
			this.textSubject.Name = "textSubject";
			this.textSubject.Size = new System.Drawing.Size(336, 20);
			this.textSubject.TabIndex = 3;
			this.textSubject.Text = "Simple MAPI.NET";
			// 
			// labelSubj
			// 
			this.labelSubj.Location = new System.Drawing.Point(7, 107);
			this.labelSubj.Name = "labelSubj";
			this.labelSubj.Size = new System.Drawing.Size(49, 16);
			this.labelSubj.TabIndex = 1;
			this.labelSubj.Text = "Subject";
			// 
			// textMail
			// 
			this.textMail.Anchor = (((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right);
			this.textMail.Font = new System.Drawing.Font("Courier New", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.textMail.HideSelection = false;
			this.textMail.Location = new System.Drawing.Point(0, 128);
			this.textMail.Multiline = true;
			this.textMail.Name = "textMail";
			this.textMail.ScrollBars = System.Windows.Forms.ScrollBars.Both;
			this.textMail.Size = new System.Drawing.Size(408, 200);
			this.textMail.TabIndex = 4;
			this.textMail.Text = "Thanks for this great sample!";
			// 
			// textCC
			// 
			this.textCC.Anchor = ((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right);
			this.textCC.Location = new System.Drawing.Point(64, 40);
			this.textCC.MaxLength = 5000;
			this.textCC.Name = "textCC";
			this.textCC.Size = new System.Drawing.Size(336, 20);
			this.textCC.TabIndex = 7;
			this.textCC.Text = "";
			// 
			// buttonSend
			// 
			this.buttonSend.Anchor = (System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right);
			this.buttonSend.Location = new System.Drawing.Point(320, 72);
			this.buttonSend.Name = "buttonSend";
			this.buttonSend.Size = new System.Drawing.Size(80, 24);
			this.buttonSend.TabIndex = 5;
			this.buttonSend.Text = "Send!";
			this.buttonSend.Click += new System.EventHandler(this.buttonSend_Click);
			// 
			// comboAttachm
			// 
			this.comboAttachm.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.comboAttachm.DropDownWidth = 300;
			this.comboAttachm.Enabled = false;
			this.comboAttachm.Location = new System.Drawing.Point(64, 72);
			this.comboAttachm.Name = "comboAttachm";
			this.comboAttachm.Size = new System.Drawing.Size(208, 21);
			this.comboAttachm.TabIndex = 9;
			// 
			// buttonAttach
			// 
			this.buttonAttach.Location = new System.Drawing.Point(8, 71);
			this.buttonAttach.Name = "buttonAttach";
			this.buttonAttach.Size = new System.Drawing.Size(56, 24);
			this.buttonAttach.TabIndex = 8;
			this.buttonAttach.Text = "Attach...";
			this.buttonAttach.Click += new System.EventHandler(this.buttonAttach_Click);
			// 
			// buttonAddrTO
			// 
			this.buttonAddrTO.Location = new System.Drawing.Point(8, 5);
			this.buttonAddrTO.Name = "buttonAddrTO";
			this.buttonAddrTO.Size = new System.Drawing.Size(56, 24);
			this.buttonAddrTO.TabIndex = 1;
			this.buttonAddrTO.Text = "TO...";
			this.buttonAddrTO.Click += new System.EventHandler(this.buttonAddrTO_Click);
			// 
			// buttonAddrCC
			// 
			this.buttonAddrCC.Location = new System.Drawing.Point(8, 38);
			this.buttonAddrCC.Name = "buttonAddrCC";
			this.buttonAddrCC.Size = new System.Drawing.Size(56, 24);
			this.buttonAddrCC.TabIndex = 6;
			this.buttonAddrCC.Text = "CC...";
			this.buttonAddrCC.Click += new System.EventHandler(this.buttonAddrCC_Click);
			// 
			// SendForm
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(408, 325);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.comboAttachm,
																		  this.buttonSend,
																		  this.textMail,
																		  this.textTO,
																		  this.textSubject,
																		  this.labelSubj,
																		  this.textCC,
																		  this.buttonAttach,
																		  this.buttonAddrTO,
																		  this.buttonAddrCC});
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.MinimumSize = new System.Drawing.Size(400, 280);
			this.Name = "SendForm";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "Send Mail";
			this.ResumeLayout(false);

		}
		#endregion

		private void buttonSend_Click( object sender, System.EventArgs e )
			{
			if( (textTO.Text == null) || (textSubject.Text == null) )
				return;
			if( (textTO.Text.Length == 0) || (textSubject.Text.Length == 0) )
				return;
			ma.AddRecip( textTO.Text, null, false );
			if( textCC.Text != null )
				{
				if( textCC.Text.Length > 0 )
					ma.AddRecip( textCC.Text, null, true );
				}

			if( ! ma.Send( textSubject.Text, textMail.Text ) )
				MessageBox.Show( this, "MAPISendMail failed! " + ma.Error(), "Send Mail", MessageBoxButtons.OK, MessageBoxIcon.Warning );

			ma.Reset();
			this.Close();
			}

		private void buttonAttach_Click( object sender, System.EventArgs e )
			{
			OpenFileDialog af = new OpenFileDialog();
			af.Title = "Attach File";
			af.Filter = "Any File (*.*)|*.*";
        
			if( af.ShowDialog() != DialogResult.OK )
				return;

			comboAttachm.Enabled = true;
			int n = comboAttachm.Items.Add( af.FileName );
			comboAttachm.SelectedIndex = n;
			ma.Attach( af.FileName );
			}

		private void buttonAddrTO_Click(object sender, System.EventArgs e)
			{
			string name; string addr;
			if( ma.SingleAddress( null, out name, out addr ) )
				textTO.Text = name;
			}

		private void buttonAddrCC_Click(object sender, System.EventArgs e)
			{
			string name; string addr;
			if( ma.SingleAddress( "CC", out name, out addr ) )
				textCC.Text = name;
			}
	}
}
