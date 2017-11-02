/* ----------------------------------- File Header -------------------------------------------*
	File				:	PluginManager.cs
	Project Code		:	
	Author	    		:	Tom Thomas
	Created On	    	:	15/03/2006 11:20:47 AM
	Last Modified	   	:	15/03/2006 11:20:47 AM
----------------------------------------------------------------------------------------------*
	Type				:	c sharp file
	Description			:   This class handles all addin related activities.
	Developer's Note	:	The code is not production level .. can expect bugs..
	Bugs				:	
	See Also			:	
	Revision History	:	
	Traceability        :	
	Necessary Files		:	
---------------------------------------------------------------------------------------------*/

using System;
using System.IO;
using System.Configuration ;
using System.Reflection;
using System.Collections;
using System.Windows.Forms;
using System.Xml;
using System.Drawing;

namespace ProjectFramework
{
	/// <summary>
	/// This class handles all the plugin related activities their loading, saving etc 
	/// </summary>
	/// 
	public struct AddinCommadInfo
	{
		public ArrayList MenuStringsArray; //All Menus before coming to the leaf node
		public int iCommandID;
		public string strMenuString;
		public string strHelpString;
		public string strToolTip;
		public string strFunction;
		public string strShortCutKey;
		public string strToolbarIndexName;
		public int iSeparator;
		public MenuItem Menu;
		void Initialize()
		{
			iCommandID=-1;
			strMenuString="";
			strHelpString="";
			strToolTip="";
			strFunction="";
			strShortCutKey="";
			strToolbarIndexName="";
			iSeparator=-1;
			Menu= new MenuItem();
			
		}
	};
	public struct AddinInfo
	{
		public string strAddinName;
		public string strAddinDllName;
		public string strAddinVersion;
		public object lInstanceHandle;
		public object lToobarRes;
		public bool bLoadAddin;
		public int lToolbarButtonCount;
		public Assembly AddinAssembly;
		public Type AddinType; 
		public IProjectFrameworkAddin FrameworkAddin;
		public string strAddinInterfaceName;
		public ArrayList AddinCommadInfoArray;

		void Initialize()
		{

		}
		
	};
	public class PluginManager
	{
		public AddinInfo[] AddinInfoArray;
		public AddinProjectFramework ProjectFramework;
		public IProjectFrameworkApp m_FrameworkApp;
		public bool m_bLoadAddinsOnStartup;
		private string m_strXMLFileName;
        public PluginManager()
		{
			// 
			// TODO: Add constructor logic here
			//
			m_bLoadAddinsOnStartup=false;
			m_strXMLFileName="";
		}

		public bool LoadPluginDetailsFromXML(string strXMLFileName)
		{
			m_strXMLFileName=strXMLFileName;
			//Load all the addin name and interface from the configuration file
			//Now hard code it for the time being
			try
			{
				XmlDocument AddinSettingsXML = new XmlDocument();
				AddinSettingsXML.Load(strXMLFileName);
				
				XmlNode LoadOnStartupNode=AddinSettingsXML.FirstChild.NextSibling;
				string strLoadOnStart=LoadOnStartupNode["LoadAddinsOnStartup"].InnerText;
				if(strLoadOnStart=="1")
				{
					m_bLoadAddinsOnStartup=true;
				}
				else
				{
					m_bLoadAddinsOnStartup=false;
				}
				XmlNodeList LeafNodes = AddinSettingsXML.GetElementsByTagName("Addin");

				AddinInfoArray= new AddinInfo[LeafNodes.Count];

				for(int i=0;i<LeafNodes.Count;i++)
				{
					XmlNodeList ValueList=LeafNodes[i].ChildNodes;
					for(int j=0;j<ValueList.Count;j++)
					{
						if(ValueList[j].Name=="DllName")
						{
							AddinInfoArray[i].strAddinDllName= ValueList[j].InnerText;
	
						}
						else if(ValueList[j].Name=="AddinInterfaceName")
						{
							AddinInfoArray[i].strAddinInterfaceName= ValueList[j].InnerText;
							
						}
						else if(ValueList[j].Name=="LoadAddin")
						{
							AddinInfoArray[i].bLoadAddin= Convert.ToBoolean(Convert.ToInt16(ValueList[j].InnerText));
							
						}
					}
				}
				
			}
			catch(Exception Ex)
			{
				MessageBox.Show(Ex.Message);
				//Create a dummy array then exit
				AddinInfoArray= new AddinInfo[0];
				return false;
			}
			return m_bLoadAddinsOnStartup;
		
		}
		public bool SavePluginDetailsToXML()
		{
			try
			{
				string strData;
				XmlDocument AddinSettingsXML = new XmlDocument();
				AddinSettingsXML.Load(m_strXMLFileName);
				
				XmlNode LoadOnStartupNode=AddinSettingsXML.FirstChild.NextSibling;
				if(m_bLoadAddinsOnStartup)
				{
					strData ="1";
				}
				else
				{
					strData ="0";
				}
				LoadOnStartupNode["LoadAddinsOnStartup"].InnerText=strData;
#if _BUG_FIXED_ON_THIS_AREA
				#region				
				XmlNodeList LeafNodes = AddinSettingsXML.GetElementsByTagName("Addin");
				for(int i=0;i<AddinInfoArray.Length;i++)
				{
					if(AddinInfoArray[i].bLoadAddin)
					{
						strData ="1";
					}
					else
					{
						strData ="0";
					}
					XmlNode Node=LeafNodes.Item(0);
					Node["LoadAddin"].InnerText=strData;
					int z=0;

				}
				#endregion
#endif
				AddinSettingsXML.Save(m_strXMLFileName);
				
			}
			catch(Exception Ex)
			{
				MessageBox.Show(Ex.Message);
				return false;
			}
			return m_bLoadAddinsOnStartup;
		}

		/// <summary>
		/// Loads all the assembly info from the addin dll
		/// </summary>
		public bool LoadAddAddinAssemblyInfo(string strAddinFolderName)
		{
			try
			{
				string path = Application.StartupPath;
				path+="\\"+strAddinFolderName;
				string[] pluginFiles = Directory.GetFiles(path, "*.DLL");
				
				//((IProjectFrameworkApp)ProjectFrameworkApp.m_FrameworkApp)
				//Get the instance of IProjectFrameworkApp
				//(IProjectFrameworkApp)FrameworkApp
				for(int i= 0; i<pluginFiles.Length; i++)
				{
					string OriginalPath = pluginFiles[i];
					string strDllName=pluginFiles[i].Substring(pluginFiles[i].LastIndexOf("\\")+1,pluginFiles[i].Length- pluginFiles[i].LastIndexOf("\\")-1); 
					//Get the interface name corresponding to the file name
					string strInterfaceName=GetMainInterfaceName(strDllName);
					if(strInterfaceName!="")
					{
						int iIndex=GetAddinIndex(strDllName);
						//Proceed to load this dll using the interface name.
						AddinInfoArray[iIndex].AddinAssembly=Assembly.LoadFrom(OriginalPath);
						AddinInfoArray[iIndex].AddinType=AddinInfoArray[iIndex].AddinAssembly.GetType(strInterfaceName);
						//Check the IProjectFrameworkAddin is implemented
						AddinInfoArray[iIndex].FrameworkAddin=(IProjectFrameworkAddin)Activator.CreateInstance(AddinInfoArray[iIndex].AddinType);
						//If so call its initializing methods
						AddinInfoArray[iIndex].FrameworkAddin.InitializeAddin(ref m_FrameworkApp,iIndex,true);
					}
					else
					{
						//Discard this dll as this is not configured in configuration file
					}

				}

			}
			catch(Exception Ex)
			{
				MessageBox.Show(Ex.Message);
				return false;
			}
			return true;
		}
		public bool UnloadAllAddins()
		{
			return true;
		}
		public bool InvokeAddinMember(int iAddinIndex, string strFunctionName)
		{
			try
			{
				
				Type tObjectType = AddinInfoArray[iAddinIndex].AddinAssembly.GetType(AddinInfoArray[iAddinIndex].strAddinInterfaceName.Trim());
				Object result = tObjectType.InvokeMember(strFunctionName.Trim(), 
					BindingFlags.CreateInstance, null, null, null);
				tObjectType.InvokeMember(strFunctionName.Trim(), 
					BindingFlags.InvokeMethod, null, result, null);
			
			}
			catch(Exception Ex)
			{
				MessageBox.Show(Ex.Message);
				return false;
			}
			return true;
		}

		public string GetMainInterfaceName(string strDllName)
		{
			foreach(AddinInfo info in AddinInfoArray)
			{
				if(info.strAddinDllName==strDllName)
				{
					return info.strAddinInterfaceName; 
				}
			}
			return "";
		}

		public int GetAddinIndex(string strDllName)
		{
			int iIndex=0;
            foreach(AddinInfo info in AddinInfoArray)
			{
				if(info.strAddinDllName==strDllName)
				{
					return iIndex; 
				}
				iIndex++;
			}
			return -1;
		}

		public bool InvokeMenu(ref object AddinMenuItem)
		{
			MenuItem Item=(MenuItem)AddinMenuItem;
			for(int i=0;i<AddinInfoArray.Length;i++)
			{
				for(int j=0;j<AddinInfoArray[i].AddinCommadInfoArray.Count;j++)
				{
					if(((AddinCommadInfo)AddinInfoArray[i].AddinCommadInfoArray[j]).Menu==Item)
					{
						return InvokeAddinMember(i,((AddinCommadInfo)AddinInfoArray[i].AddinCommadInfoArray[j]).strFunction);
					}
				}
			}
			return false;
		}
		
		public bool UpdateAddinMenuStatus(int iAddinIndex, bool bEnable)
		{
			AddinInfoArray[iAddinIndex].bLoadAddin=bEnable;
			for(int j=0;j<AddinInfoArray[iAddinIndex].AddinCommadInfoArray.Count;j++)
			{
				((AddinCommadInfo)AddinInfoArray[iAddinIndex].AddinCommadInfoArray[j]).Menu.Enabled=bEnable;
			}
			//Update the toolbars if any
			for(int i=0;i<ProjectFramework.m_AddinToolbarArray.Count;i++)
			{
				if(((ToolBar)ProjectFramework.m_AddinToolbarArray[i]).Name==AddinInfoArray[iAddinIndex].strAddinName)
				{
					((ToolBar)ProjectFramework.m_AddinToolbarArray[i]).Enabled=bEnable;
				}
			}
			return true;
		}

		public bool LoadAddinMenus()
		{
			//Add each Addin menu in appropriate menu item
			
			for(int i=0;i<AddinInfoArray.Length;i++)
			{
				//Iterate through each addin command array
				for(int j=0;j<AddinInfoArray[i].AddinCommadInfoArray.Count;j++)
				{
					//Get each comand info 
					AddinCommadInfo CommadInfo=(AddinCommadInfo)AddinInfoArray[i].AddinCommadInfoArray[j];
					//Check whether the particular main menu and sub menu is there
					//If not there then create and add it before help menu
					
					//Get the main menu
					string strMainMenuString=(string)CommadInfo.MenuStringsArray[CommadInfo.MenuStringsArray.Count-1];
					MenuItem MainMenuItem=null;
					GetMainMenuItem(strMainMenuString,out MainMenuItem);
					//if main menu is null then create the menu
					if(MainMenuItem==null)
					{
						MainMenuItem= new MenuItem(strMainMenuString);
						ProjectFramework.mainProjectMenu.MenuItems.Add(MainMenuItem);
																		
					}
					//Now iterate through all sub menus 
					//if the sub menu is not there then create it
					MenuItem SubMenuItem=null;
					MenuItem ParentMenu=MainMenuItem;
					for(int k=CommadInfo.MenuStringsArray.Count-2;k>=0;k--)
					{
						GetMenuItem((string)CommadInfo.MenuStringsArray[k],ref ParentMenu,out SubMenuItem);
						if(SubMenuItem==null)
						{
							SubMenuItem= new MenuItem((string)CommadInfo.MenuStringsArray[k]);
							ParentMenu.MenuItems.Add(SubMenuItem);
						}
						ParentMenu=SubMenuItem;
					}
					//Got the sub menu .. now add the leaf menu
					// With appropriate menu handler
					CommadInfo.Menu= new MenuItem(CommadInfo.strMenuString);
					AddinInfoArray[i].AddinCommadInfoArray[j]=CommadInfo;
					CommadInfo.Menu.Click+= new System.EventHandler(ProjectFramework.menuItemPlugin_Click);
					CommadInfo.Menu.Select+=new System.EventHandler(ProjectFramework.menuItem_Select);
					//CommadInfo.Menu.Text=CommadInfo.Menu.Text+" "+CommadInfo.strShortCutKey;
					string strChortCutKey="";
					
					if(CommadInfo.strShortCutKey!=null)
					{
						try
						{
							strChortCutKey=CommadInfo.strShortCutKey.Replace("+","");
							Shortcut ShortcutKey=(Shortcut)Enum.Parse(typeof(Shortcut),strChortCutKey);
							CommadInfo.Menu.Shortcut=ShortcutKey;
						}
						catch(Exception ex)
						{
							MessageBox.Show(ex.Message);
						}
								 
					}
					
					ParentMenu.MenuItems.Add(CommadInfo.Menu);
				}
				
			}
			return true;
			
		}
		public bool LoadAddinToolbars()
		{
			//Load all the Images of the plugins in the right order
			int iExistingImageCount=ProjectFramework.imageListMain.Images.Count;
			int iPluginImageIndex=iExistingImageCount;
			//Now add all the images from the plugins in the order
			//in which it is retrieved
			for(int i=0;i<AddinInfoArray.Length;i++)
			{
				for(int j=0;j<AddinInfoArray[i].AddinCommadInfoArray.Count;j++)
				{
					string ToolbarIconName=((AddinCommadInfo)AddinInfoArray[i].AddinCommadInfoArray[j]).strToolbarIndexName;
					if(ToolbarIconName!=null)
					{
						string strAssembleyIconName=AddinInfoArray[i].AddinAssembly.GetName().Name+"."+ToolbarIconName;
						Stream IconStream = AddinInfoArray[i].AddinAssembly.GetManifestResourceStream(strAssembleyIconName);
						if(IconStream!=null)
						{
							Icon Addinicon = new Icon( IconStream );
							ProjectFramework.imageListMain.Images.Add(Addinicon);
						}
					}
				}
			}

			for(int i=0;i<AddinInfoArray.Length;i++)
			{
				if(AddinInfoArray[i].lToolbarButtonCount<1)
				{
					continue;
				}
				//Create the toolbar
				ToolBar toolBarPlugin= new ToolBar();
				toolBarPlugin.DropDownArrows = true;
				toolBarPlugin.AutoSize=true;
				toolBarPlugin.Dock=DockStyle.Top; 
				toolBarPlugin.ImageList = ProjectFramework.imageListMain;
				toolBarPlugin.Location = new System.Drawing.Point(0, 0);
				toolBarPlugin.Name = AddinInfoArray[i].strAddinName;
				toolBarPlugin.ShowToolTips = true;
				toolBarPlugin.Size = new System.Drawing.Size(720, 28);
				toolBarPlugin.TabIndex = 0;
				toolBarPlugin.ButtonClick += new System.Windows.Forms.ToolBarButtonClickEventHandler(ProjectFramework.toolBarPlugin_ButtonClick);
				//Add the button too
				for(int j=0;j<AddinInfoArray[i].AddinCommadInfoArray.Count;j++)
				{
					string ToolbarIconName=((AddinCommadInfo)AddinInfoArray[i].AddinCommadInfoArray[j]).strToolbarIndexName;
					if(ToolbarIconName!=null)
					{
						ToolBarButton TBBuutton = new ToolBarButton();
						TBBuutton.ImageIndex = iPluginImageIndex;
						TBBuutton.ToolTipText = ((AddinCommadInfo)AddinInfoArray[i].AddinCommadInfoArray[j]).strToolTip;
						TBBuutton.Tag=((AddinCommadInfo)AddinInfoArray[i].AddinCommadInfoArray[j]).strFunction;

						toolBarPlugin.Buttons.Add(TBBuutton);
						iPluginImageIndex++;
					}
				}
				ProjectFramework.Controls.Add(toolBarPlugin);
				ProjectFramework.m_AddinToolbarArray.Add((object)toolBarPlugin);
			}
			return true;
		}
		public bool LoadAddinToolbarMenus()
		{
			foreach(ToolBar ToolBarInfo in ProjectFramework.m_AddinToolbarArray)
			{
				//
				MenuItem MainMenuItem= new MenuItem(ToolBarInfo.Name);
				MainMenuItem.Checked=true;
				MainMenuItem.Click+= new System.EventHandler(ProjectFramework.addinToolbarMenuItem_Click);
				ProjectFramework.menuItemToolbar.MenuItems.Add(MainMenuItem);
			}
			return true;
		}

		public void GetMainMenuItem(string strMenuName, out MenuItem Item)
		{
			Item=null;
			for(int i=0;i<ProjectFramework.mainProjectMenu.MenuItems.Count;i++)
			{
				if(ProjectFramework.mainProjectMenu.MenuItems[i].Text==strMenuName)
				{
					Item=ProjectFramework.mainProjectMenu.MenuItems[i];
				}
			}
		}
		public void GetMenuItem(string strMenuName,ref MenuItem ItemParent, out MenuItem Item)
		{
			Item=null;
			for(int i=0;i<ItemParent.MenuItems.Count ;i++)
			{
				if(ItemParent.MenuItems[i].Text==strMenuName)
				{
					Item=ItemParent.MenuItems[i];
				}
			}
		}


		/// <summary>
		/// Set the project Framework app
		/// </summary>
		
		/// <summary>
		/// Get The Addin Information given the index
		/// </summary>

	}
}
