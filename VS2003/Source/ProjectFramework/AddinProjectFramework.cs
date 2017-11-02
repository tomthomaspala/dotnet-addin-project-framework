using System;
using System.Reflection;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
namespace ProjectFramework
{

	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class AddinProjectFramework : System.Windows.Forms.Form
	{
		
		private System.Windows.Forms.MenuItem menuItem1;
		private System.Windows.Forms.MenuItem menuItem3;
		private System.Windows.Forms.MenuItem menuItem6;
		private System.Windows.Forms.MenuItem menuItemExit;
		public System.Windows.Forms.MenuItem menuItemToolbar;
		private System.Windows.Forms.MenuItem menuItemStatusBar;
		private System.Windows.Forms.MenuItem menuItemAddinSettings;
		private System.Windows.Forms.MenuItem menuItemAbout;
		public System.Windows.Forms.MainMenu mainProjectMenu;
		private System.Windows.Forms.ToolBar toolBarMain;
		private System.Windows.Forms.ToolBarButton toolBarButton;
		private System.Windows.Forms.ToolBarButton toolBarButtonSettings;
		public System.Windows.Forms.ImageList imageListMain;
		private System.Windows.Forms.ToolBarButton toolBarButtonAbout;
		private System.Windows.Forms.ContextMenu contextMenu1;
		private System.Windows.Forms.MenuItem menuItem2;
		private System.ComponentModel.IContainer components;
		private bool m_bLoadAllAddins;
		private System.Windows.Forms.StatusBar PFStatusBar;
		private System.Windows.Forms.StatusBarPanel statusBarMenuText;
		private System.Windows.Forms.StatusBarPanel statusBarAddin;
		public PluginManager m_PluginManager;
		public const string PF_CONFIG_FILE_NAME= "ProjectFramework.exe.config";
		private System.Windows.Forms.Label labelText;
		public const string PF_ADDIN_FOLDER_NAME= "Addins";
		public ProjectFrameworkApp m_FrameworkApp;
		public string m_strPluginFile;
		private System.Windows.Forms.MenuItem menuItem4;
		private System.Windows.Forms.MenuItem menuItemHelp;
		public ArrayList m_AddinToolbarArray;
		public AddinProjectFramework()
		{
			//
			// Required for Windows Form Designer support
			//
			m_PluginManager= new PluginManager();
			m_FrameworkApp = new  ProjectFrameworkApp();
			m_AddinToolbarArray= new ArrayList();
			System.Configuration.AppSettingsReader configurationAppSettings = new System.Configuration.AppSettingsReader();
			m_strPluginFile=((string)(configurationAppSettings.GetValue("PluginFile", typeof(string))));
			InitializeComponent();
			
			m_PluginManager.ProjectFramework=this;
			m_FrameworkApp.ProjectFramework=this;
			
			//Set the IProjectFrameworkApp reference

			m_PluginManager.m_FrameworkApp=(IProjectFrameworkApp)m_FrameworkApp; 
			
			//Load all addins
			
			if(CheckAddinSettings())
			{
				LoadAllAddins(m_bLoadAllAddins);
				m_PluginManager.LoadAddinMenus();
				m_PluginManager.LoadAddinToolbars();
				m_PluginManager.LoadAddinToolbarMenus();
				//Remove the about box from the Middle and add it towards the end.
				mainProjectMenu.MenuItems.RemoveAt(3); //Help menu is here
				mainProjectMenu.MenuItems.Add(menuItemHelp);
			}
			
			
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(AddinProjectFramework));
			System.Configuration.AppSettingsReader configurationAppSettings = new System.Configuration.AppSettingsReader();
			this.mainProjectMenu = new System.Windows.Forms.MainMenu();
			this.menuItem1 = new System.Windows.Forms.MenuItem();
			this.menuItemExit = new System.Windows.Forms.MenuItem();
			this.menuItem3 = new System.Windows.Forms.MenuItem();
			this.menuItemToolbar = new System.Windows.Forms.MenuItem();
			this.menuItem4 = new System.Windows.Forms.MenuItem();
			this.menuItemStatusBar = new System.Windows.Forms.MenuItem();
			this.menuItem6 = new System.Windows.Forms.MenuItem();
			this.menuItemAddinSettings = new System.Windows.Forms.MenuItem();
			this.menuItemHelp = new System.Windows.Forms.MenuItem();
			this.menuItemAbout = new System.Windows.Forms.MenuItem();
			this.toolBarMain = new System.Windows.Forms.ToolBar();
			this.toolBarButton = new System.Windows.Forms.ToolBarButton();
			this.toolBarButtonSettings = new System.Windows.Forms.ToolBarButton();
			this.toolBarButtonAbout = new System.Windows.Forms.ToolBarButton();
			this.imageListMain = new System.Windows.Forms.ImageList(this.components);
			this.contextMenu1 = new System.Windows.Forms.ContextMenu();
			this.menuItem2 = new System.Windows.Forms.MenuItem();
			this.labelText = new System.Windows.Forms.Label();
			this.PFStatusBar = new System.Windows.Forms.StatusBar();
			this.statusBarMenuText = new System.Windows.Forms.StatusBarPanel();
			this.statusBarAddin = new System.Windows.Forms.StatusBarPanel();
			((System.ComponentModel.ISupportInitialize)(this.statusBarMenuText)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.statusBarAddin)).BeginInit();
			this.SuspendLayout();
			// 
			// mainProjectMenu
			// 
			this.mainProjectMenu.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																							this.menuItem1,
																							this.menuItem3,
																							this.menuItem6,
																							this.menuItemHelp});
			// 
			// menuItem1
			// 
			this.menuItem1.Index = 0;
			this.menuItem1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																					  this.menuItemExit});
			this.menuItem1.Text = "File";
			// 
			// menuItemExit
			// 
			this.menuItemExit.Index = 0;
			this.menuItemExit.Shortcut = System.Windows.Forms.Shortcut.Alt4;
			this.menuItemExit.Text = "Exit";
			this.menuItemExit.Click += new System.EventHandler(this.menuItemExit_Click);
			this.menuItemExit.Select += new System.EventHandler(this.menuItemExit_Select);
			// 
			// menuItem3
			// 
			this.menuItem3.Index = 1;
			this.menuItem3.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																					  this.menuItemToolbar,
																					  this.menuItemStatusBar});
			this.menuItem3.Text = "View";
			// 
			// menuItemToolbar
			// 
			this.menuItemToolbar.Index = 0;
			this.menuItemToolbar.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																							this.menuItem4});
			this.menuItemToolbar.Text = "&Toolbar";
			// 
			// menuItem4
			// 
			this.menuItem4.Checked = true;
			this.menuItem4.Index = 0;
			this.menuItem4.Text = "Main Toolbar";
			this.menuItem4.Click += new System.EventHandler(this.menuItem4_Click);
			// 
			// menuItemStatusBar
			// 
			this.menuItemStatusBar.Checked = true;
			this.menuItemStatusBar.Index = 1;
			this.menuItemStatusBar.Text = "Status Bar";
			this.menuItemStatusBar.Click += new System.EventHandler(this.menuItemStatusBar_Click);
			// 
			// menuItem6
			// 
			this.menuItem6.Index = 2;
			this.menuItem6.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																					  this.menuItemAddinSettings});
			this.menuItem6.Text = "Settings";
			// 
			// menuItemAddinSettings
			// 
			this.menuItemAddinSettings.Index = 0;
			this.menuItemAddinSettings.Text = "Addin Settings";
			this.menuItemAddinSettings.Click += new System.EventHandler(this.menuItemAddinSettings_Click);
			// 
			// menuItemHelp
			// 
			this.menuItemHelp.Index = 3;
			this.menuItemHelp.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																						 this.menuItemAbout});
			this.menuItemHelp.Text = "Help";
			// 
			// menuItemAbout
			// 
			this.menuItemAbout.Index = 0;
			this.menuItemAbout.Text = "About Addin Project Framework";
			this.menuItemAbout.Click += new System.EventHandler(this.menuItemAbout_Click);
			// 
			// toolBarMain
			// 
			this.toolBarMain.Buttons.AddRange(new System.Windows.Forms.ToolBarButton[] {
																						   this.toolBarButton,
																						   this.toolBarButtonSettings,
																						   this.toolBarButtonAbout});
			this.toolBarMain.DropDownArrows = true;
			this.toolBarMain.ImageList = this.imageListMain;
			this.toolBarMain.Location = new System.Drawing.Point(0, 0);
			this.toolBarMain.Name = "toolBarMain";
			this.toolBarMain.ShowToolTips = true;
			this.toolBarMain.Size = new System.Drawing.Size(720, 44);
			this.toolBarMain.TabIndex = 0;
			this.toolBarMain.ButtonClick += new System.Windows.Forms.ToolBarButtonClickEventHandler(this.toolBarMain_ButtonClick);
			// 
			// toolBarButton
			// 
			this.toolBarButton.Style = System.Windows.Forms.ToolBarButtonStyle.Separator;
			// 
			// toolBarButtonSettings
			// 
			this.toolBarButtonSettings.ImageIndex = 0;
			this.toolBarButtonSettings.ToolTipText = "Addin Settings";
			// 
			// toolBarButtonAbout
			// 
			this.toolBarButtonAbout.ImageIndex = 1;
			this.toolBarButtonAbout.ToolTipText = "About";
			// 
			// imageListMain
			// 
			this.imageListMain.ImageSize = new System.Drawing.Size(32, 32);
			this.imageListMain.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageListMain.ImageStream")));
			this.imageListMain.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// contextMenu1
			// 
			this.contextMenu1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																						 this.menuItem2});
			// 
			// menuItem2
			// 
			this.menuItem2.Index = 0;
			this.menuItem2.Text = "Exit";
			// 
			// labelText
			// 
			this.labelText.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelText.ForeColor = System.Drawing.SystemColors.ActiveCaption;
			this.labelText.Location = new System.Drawing.Point(112, 168);
			this.labelText.Name = "labelText";
			this.labelText.Size = new System.Drawing.Size(480, 64);
			this.labelText.TabIndex = 1;
			this.labelText.Text = "Dotnet Based Addin/ Plugin Framework With Dynamic Toolbars and Menus";
			// 
			// PFStatusBar
			// 
			this.PFStatusBar.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.PFStatusBar.Location = new System.Drawing.Point(0, 387);
			this.PFStatusBar.Name = "PFStatusBar";
			this.PFStatusBar.Panels.AddRange(new System.Windows.Forms.StatusBarPanel[] {
																						   this.statusBarMenuText,
																						   this.statusBarAddin});
			this.PFStatusBar.ShowPanels = true;
			this.PFStatusBar.Size = new System.Drawing.Size(720, 22);
			this.PFStatusBar.TabIndex = 2;
			this.PFStatusBar.Text = "status";
			// 
			// statusBarMenuText
			// 
			this.statusBarMenuText.Text = ((string)(configurationAppSettings.GetValue("statusBarMenuText.Text", typeof(string))));
			this.statusBarMenuText.Width = 200;
			// 
			// statusBarAddin
			// 
			this.statusBarAddin.Text = "Addin Status";
			this.statusBarAddin.Width = 200;
			// 
			// AddinProjectFramework
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.BackColor = System.Drawing.SystemColors.Info;
			this.ClientSize = new System.Drawing.Size(720, 409);
			this.Controls.Add(this.PFStatusBar);
			this.Controls.Add(this.labelText);
			this.Controls.Add(this.toolBarMain);
			this.Menu = this.mainProjectMenu;
			this.Name = "AddinProjectFramework";
			this.Text = "Addin Project Framework";
			this.Load += new System.EventHandler(this.AddinProjectFramework_Load);
			((System.ComponentModel.ISupportInitialize)(this.statusBarMenuText)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.statusBarAddin)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new AddinProjectFramework());
		}

		private void menuItemExit_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void menuItemAddinSettings_Click(object sender, System.EventArgs e)
		{
			AddinSettings Settings= new AddinSettings();
			Settings.ProjectFramework=this;
			Settings.ShowDialog();
		}

		private void menuItemAbout_Click(object sender, System.EventArgs e)
		{
			AboutBox About = new AboutBox();
			About.ShowDialog(); 
		}

		private void toolBarMain_ButtonClick(object sender, System.Windows.Forms.ToolBarButtonClickEventArgs e)
		{
			if(e.Button.ImageIndex==0)
			{
				AddinSettings Settings= new AddinSettings();
				Settings.ProjectFramework=this;
				Settings.ShowDialog();
				
			}
			else if(e.Button.ImageIndex==1)
			{
				AboutBox About = new AboutBox();
				About.ShowDialog(); 
			}
		}
		public void toolBarPlugin_ButtonClick(object sender, System.Windows.Forms.ToolBarButtonClickEventArgs e)
		{
			ToolBar PluginToolbar=(ToolBar)sender;
			string strFunctionName=(string)e.Button.Tag;
			//Now check from which toolbar and to which addin
			int iAddinIndex=0;
			foreach(ToolBar ToolBarInfo in m_AddinToolbarArray)
			{
				if(ToolBarInfo==PluginToolbar)
				{
					if(m_PluginManager.AddinInfoArray[iAddinIndex].bLoadAddin)
					{
						m_PluginManager.InvokeAddinMember(iAddinIndex,strFunctionName);
						return;
					}
				}
				iAddinIndex++;
			}
		}
		/// <summary>
		/// This function will load all addins based on the flag bLoadAllAddins
		/// </summary>
		public bool LoadAllAddins(bool bLoadAllAddins)
		{
			if(bLoadAllAddins)
			{
				return LoadAddAddinAssemblyInfo();
			}
			else
			{
				return false;
			}
			
		}
		private bool LoadAddAddinAssemblyInfo()
		{
			return m_PluginManager.LoadAddAddinAssemblyInfo(PF_ADDIN_FOLDER_NAME);
		}
		private bool CheckAddinSettings()
		{
			if(LoadAddinInfoFromXML())
			{
				m_bLoadAllAddins=true;
			}
			else
			{
				m_bLoadAllAddins=false;
			}
			return m_bLoadAllAddins;
		}
		private bool LoadAddinInfoFromXML()
		{
			return m_PluginManager.LoadPluginDetailsFromXML(m_strPluginFile);
		}

		public void SendMessage(string strMessage)
		{
			labelText.Text=strMessage;
		}
		public void menuItemPlugin_Click(object sender, System.EventArgs e)
		{
			bool bSuccess=m_PluginManager.InvokeMenu( ref sender);
		}
		
		private void menuItemExit_Select(object sender, System.EventArgs e)
		{
			this.statusBarMenuText.Text="Exit from the App";
		}
		public void menuItem_Select(object sender, System.EventArgs e)
		{
			this.statusBarMenuText.Text=((MenuItem)sender).Text;
		}

		private void AddinProjectFramework_Load(object sender, System.EventArgs e)
		{
			
		}

		private void menuItem4_Click(object sender, System.EventArgs e)
		{
			MenuItem ToolbarMenuItem=(MenuItem)sender;
			if(ToolbarMenuItem.Checked==true)
			{
				ToolbarMenuItem.Checked=false;
				toolBarMain.Visible=false;
			}
			else
			{
				ToolbarMenuItem.Checked=true;
				toolBarMain.Visible=true;
			}
		}
		
		public void addinToolbarMenuItem_Click(object sender, System.EventArgs e)
		{
			MenuItem AddinToolbarMenuItem=(MenuItem)sender;
			if(AddinToolbarMenuItem.Checked==true)
			{
				AddinToolbarMenuItem.Checked=false;
				
			}
			else
			{
				AddinToolbarMenuItem.Checked=true;
				
			}
			foreach(ToolBar ToolBarInfo in m_AddinToolbarArray)
			{
				if(ToolBarInfo.Name==AddinToolbarMenuItem.Text)
				{
					ToolBarInfo.Visible=AddinToolbarMenuItem.Checked;
				}
			}
		}

		private void menuItemStatusBar_Click(object sender, System.EventArgs e)
		{
			MenuItem StatusBarMenuItem=(MenuItem)sender;
			if(StatusBarMenuItem.Checked==true)
			{
				StatusBarMenuItem.Checked=false;
				PFStatusBar.Visible=false;
			}
			else
			{
				StatusBarMenuItem.Checked=true;
				PFStatusBar.Visible=true;
			}
		}
		
	}
}
