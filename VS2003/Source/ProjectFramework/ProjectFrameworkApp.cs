/* ----------------------------------- File Header -------------------------------------------*
	File				:	ProjectFrameworkApp.cs
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
using System.Windows.Forms;
using System.Xml;

namespace ProjectFramework
{
	/// <summary>
	/// 
	/// </summary>
	public class ProjectFrameworkApp : ProjectFramework.IProjectFrameworkApp
	{
		public AddinProjectFramework ProjectFramework;

		public ProjectFrameworkApp()
		{
			// 
			// TODO: Add constructor logic here
			//
		}

		public void AddCommandsInfo(string strXMLMenuInfo, long lSession,object lInstanceHandle, object lToolbarInfo)
		{
			//Store the addin Info to the data structure..
			ProjectFramework.m_PluginManager.AddinInfoArray[lSession].lInstanceHandle=lInstanceHandle;
			ProjectFramework.m_PluginManager.AddinInfoArray[lSession].lToobarRes=lToolbarInfo;
			if(!LoadMenuCommands(strXMLMenuInfo,lSession))
			{
				MessageBox.Show("Failed to load XML menu Info");
			}

		}
		public bool LoadMenuCommands(string strXMLMenuInfo,long lSession)
		{
			try
			{
				XmlDocument AddinSettingsXML = new XmlDocument();
				AddinSettingsXML.LoadXml(strXMLMenuInfo);

				XmlNode rootNode = AddinSettingsXML.FirstChild;
				//Display the contents of the child nodes.
				rootNode=rootNode.NextSibling;

				if (rootNode.HasChildNodes)
				{
					ProjectFramework.m_PluginManager.AddinInfoArray[lSession].strAddinName=rootNode["AddinName"].InnerText;
					ProjectFramework.m_PluginManager.AddinInfoArray[lSession].lToolbarButtonCount=Convert.ToInt32(rootNode["ToobarButtonCount"].InnerText);
					ProjectFramework.m_PluginManager.AddinInfoArray[lSession].strAddinVersion=rootNode["AppVer"].InnerText;
					//Get the addin Version
				}
				
				XmlNodeList LeafNodes = AddinSettingsXML.GetElementsByTagName("LeafMenu");
				//Add the commands into information array
				ProjectFramework.m_PluginManager.AddinInfoArray[lSession].AddinCommadInfoArray= new System.Collections.ArrayList();

				for(int i=0;i<LeafNodes.Count;i++)
				{
					XmlNodeList ValueList=LeafNodes[i].ChildNodes;
					AddinCommadInfo CommadInfo= new AddinCommadInfo();
					for(int j=0;j<ValueList.Count;j++)
					{
						if(ValueList[j].Name=="Name")
						{
							CommadInfo.strMenuString= ValueList[j].InnerText;
	
						}
						else if(ValueList[j].Name=="FunctionName")
						{
							CommadInfo.strFunction= ValueList[j].InnerText;
							
						}
						else if(ValueList[j].Name=="HelpString")
						{
							CommadInfo.strHelpString= ValueList[j].InnerText;
						}
						else if(ValueList[j].Name=="ToolTip")
						{
							CommadInfo.strToolTip= ValueList[j].InnerText;
						}
						else if(ValueList[j].Name=="ToolBarIndex")
						{
							CommadInfo.strToolbarIndexName= ValueList[j].InnerText;
						}
						else if(ValueList[j].Name=="ShortCutKey")
						{
							CommadInfo.strShortCutKey= ValueList[j].InnerText;
						}
						else if(ValueList[j].Name=="Separator")
						{
							CommadInfo.iSeparator= Convert.ToInt32(ValueList[j].InnerText);
						}
						else if(ValueList[j].Name=="ShortcutKeyIndex")
						{
							int iShortcutKeyIndex= Convert.ToInt32(ValueList[j].InnerText);
							if(CommadInfo.strMenuString.Length > iShortcutKeyIndex)
							{
								CommadInfo.strMenuString.Insert(iShortcutKeyIndex,"&");
							}
						}

					}
					//Get the parent node menu strings
					CommadInfo.MenuStringsArray= new System.Collections.ArrayList();

					XmlNode ParentNode=LeafNodes[i].ParentNode;
					while(ParentNode!=null)
					{
						XmlNodeList MenuValueList=ParentNode.ChildNodes;
						string strMenuName="";
						for(int j=0;j<MenuValueList.Count;j++)
						{
							if(MenuValueList[j].Name=="Name")
							{
								strMenuName= MenuValueList[j].InnerText;
	
							}
							else if(MenuValueList[j].Name=="ShortcutKeyIndex")
							{
								int iShortcutKeyIndex= Convert.ToInt32(MenuValueList[j].InnerText);
								if(strMenuName.Length > iShortcutKeyIndex)
								{
									strMenuName=strMenuName.Insert(iShortcutKeyIndex,"&");
									
								}	
							}

						}
						//Add this to menu array
						if(strMenuName!="")
						{
							CommadInfo.MenuStringsArray.Add(strMenuName);
						}
						if(ParentNode.Name=="MainMenu")
						{
							break;
						}
						ParentNode=ParentNode.ParentNode;
					}
					//Add this to the command info list
					ProjectFramework.m_PluginManager.AddinInfoArray[lSession].AddinCommadInfoArray.Add(CommadInfo); 

				}
				
			}
			catch(Exception Ex)
			{
				MessageBox.Show(Ex.Message);
				return false;
			}
			return true;
		}
		public void SendMessage(string strMessage)
		{
			ProjectFramework.SendMessage(strMessage);			
		}

	}
}
