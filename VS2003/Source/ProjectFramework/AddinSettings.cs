using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Xml;

namespace ProjectFramework
{
	/// <summary>
	/// Summary description for AddinSettings.
	/// </summary>
	public class AddinSettings : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button buttonOK;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.CheckedListBox checkedListBoxAddinSettings;
		private System.Windows.Forms.CheckBox checkBoxLoadAddins;
		public AddinProjectFramework ProjectFramework;
		public AddinSettings()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
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
			this.buttonOK = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.checkBoxLoadAddins = new System.Windows.Forms.CheckBox();
			this.checkedListBoxAddinSettings = new System.Windows.Forms.CheckedListBox();
			this.groupBox1.SuspendLayout();
			this.SuspendLayout();
			// 
			// buttonOK
			// 
			this.buttonOK.Location = new System.Drawing.Point(220, 228);
			this.buttonOK.Name = "buttonOK";
			this.buttonOK.Size = new System.Drawing.Size(92, 24);
			this.buttonOK.TabIndex = 1;
			this.buttonOK.Text = "OK";
			this.buttonOK.Click += new System.EventHandler(this.buttonOK_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.checkBoxLoadAddins);
			this.groupBox1.Controls.Add(this.checkedListBoxAddinSettings);
			this.groupBox1.Location = new System.Drawing.Point(24, 8);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(288, 216);
			this.groupBox1.TabIndex = 2;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Available Addins";
			// 
			// checkBoxLoadAddins
			// 
			this.checkBoxLoadAddins.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.checkBoxLoadAddins.ForeColor = System.Drawing.SystemColors.ActiveCaption;
			this.checkBoxLoadAddins.Location = new System.Drawing.Point(8, 176);
			this.checkBoxLoadAddins.Name = "checkBoxLoadAddins";
			this.checkBoxLoadAddins.Size = new System.Drawing.Size(272, 32);
			this.checkBoxLoadAddins.TabIndex = 2;
			this.checkBoxLoadAddins.Text = "Load all addins when starting the application";
			// 
			// checkedListBoxAddinSettings
			// 
			this.checkedListBoxAddinSettings.Location = new System.Drawing.Point(16, 24);
			this.checkedListBoxAddinSettings.Name = "checkedListBoxAddinSettings";
			this.checkedListBoxAddinSettings.Size = new System.Drawing.Size(240, 139);
			this.checkedListBoxAddinSettings.TabIndex = 1;
			this.checkedListBoxAddinSettings.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.checkedListBoxAddinSettings_ItemCheck);
			// 
			// AddinSettings
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(330, 256);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.buttonOK);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "AddinSettings";
			this.Text = "Addin Settings";
			this.Load += new System.EventHandler(this.AddinSettings_Load);
			this.groupBox1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void buttonOK_Click(object sender, System.EventArgs e)
		{
			try
			{
				for(int i=0;i<checkedListBoxAddinSettings.Items.Count;i++)
				{
					bool bCheck=Convert.ToBoolean(checkedListBoxAddinSettings.GetItemChecked(i));
					ProjectFramework.m_PluginManager.UpdateAddinMenuStatus(i,bCheck);
				}
				//Get the load all addin status
				ProjectFramework.m_PluginManager.m_bLoadAddinsOnStartup=checkBoxLoadAddins.Checked;
				ProjectFramework.m_PluginManager.SavePluginDetailsToXML();
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
			this.Close();
		}

		private void AddinSettings_Load(object sender, System.EventArgs e)
		{
			if(ProjectFramework.m_PluginManager.m_bLoadAddinsOnStartup)
			{
				for(int i=0;i<ProjectFramework.m_PluginManager.AddinInfoArray.Length;i++)
				{
					if(ProjectFramework.m_PluginManager.AddinInfoArray[i].strAddinName!=null)
					{
						checkedListBoxAddinSettings.Items.Add(ProjectFramework.m_PluginManager.AddinInfoArray[i].strAddinName);  
						checkedListBoxAddinSettings.SetItemChecked(i,ProjectFramework.m_PluginManager.AddinInfoArray[i].bLoadAddin);
					}
				}
			}
			checkBoxLoadAddins.Checked=ProjectFramework.m_PluginManager.m_bLoadAddinsOnStartup;
		}

		private void checkedListBoxAddinSettings_ItemCheck(object sender, System.Windows.Forms.ItemCheckEventArgs e)
		{
			ProjectFramework.m_PluginManager.AddinInfoArray[e.Index].bLoadAddin= Convert.ToBoolean(e.NewValue);
		}
	}
}
