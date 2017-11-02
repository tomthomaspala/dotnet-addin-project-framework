using System;
using ProjectFramework;
using System.Windows.Forms;
using System.IO;
using System.Xml;
using System.Reflection;


namespace PFAddin1
{
	/// <summary>
	/// 
	/// </summary>
	public class IAddin1 : ProjectFramework.IProjectFrameworkAddin
	{
		private static IProjectFrameworkApp refPFApp;
		public IAddin1()
		{
			// 
			// TODO: Add constructor logic here
			//
		}
		
		public void InitializeAddin(ref IProjectFrameworkApp refProjectFrameworkApp, long lSession, bool bFirstLoad)
		{
			refPFApp=refProjectFrameworkApp;

			Assembly thisAssembly = Assembly.GetExecutingAssembly();
			//Read the embedded XML menu resource
			Stream rgbxml = thisAssembly.GetManifestResourceStream("Addin1.Addin1XMLMenu.xml");			
			XmlDocument doc = new XmlDocument();
			doc.Load(rgbxml);
			string strMenuXML=doc.InnerXml;
			//Pass the addin dll handle and resource object..
			refProjectFrameworkApp.AddCommandsInfo(strMenuXML,lSession,(object)this.GetType().Module,null);			
		}
		public void UnInitializeAddin(long lSession)
		{
			MessageBox.Show("UnInitializeAddin","Addin 1");
		}
		
		public void Function1()
		{
			MessageBox.Show("Function1","Addin 1");
		}

		public void Function2()
		{
			Addin1Form Form = new Addin1Form();
			Form.ShowDialog();
			refPFApp.SendMessage("Hai " + Form.strMessage);
		}

		public void Function3()
		{
			MessageBox.Show("Function3","Addin 1");
		}
		
	}
}
