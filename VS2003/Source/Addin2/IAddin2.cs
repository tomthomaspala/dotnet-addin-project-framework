using System;
using ProjectFramework;
using System.Windows.Forms;
using System.IO;
using System.Xml;
using System.Reflection;

namespace PFAddin2 
{
	/// <summary>
	/// Summary description for Class1.
	/// </summary>
	public class IAddin2 : ProjectFramework.IProjectFrameworkAddin
	{
		private static IProjectFrameworkApp refPFApp;

		public IAddin2()
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
			Stream rgbxml = thisAssembly.GetManifestResourceStream("Addin2.Addin2XMLMenu.xml");			
			XmlDocument doc = new XmlDocument();
			doc.Load(rgbxml);
			string strMenuXML=doc.InnerXml;
			//Pass the addin dll handle and resource object..
			refProjectFrameworkApp.AddCommandsInfo(strMenuXML,lSession,(object)this.GetType().Module,null);			
			
		}
		public void UnInitializeAddin(long lSession)
		{
			MessageBox.Show("UnInitializeAddin","Addin 2");
		}
		
		public void Function1()
		{
			MessageBox.Show("Function1","Addin 2");
			
		}

		public void Function2()
		{
			refPFApp.SendMessage("Hi From Bar code Addin 2");
		}

		public void Function3()
		{
			MessageBox.Show("Function3","Addin 2");
		}
	}
}
