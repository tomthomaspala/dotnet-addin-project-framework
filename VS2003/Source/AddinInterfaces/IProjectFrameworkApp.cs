using System;

namespace ProjectFramework
{
	/// <summary>
	/// The framework Interface which will pass through every addin
	/// </summary>
	public interface IProjectFrameworkApp
	{
		void AddCommandsInfo(string strXMLMenuInfo, long lSession,object lInstanceHandle, object lToolbarInfo);
		void SendMessage(string strMessage);		
	}
}
