using System;

namespace ProjectFramework
{
	/// <summary>
	/// 
	/// </summary>
	public interface IProjectFrameworkAddin 
	{

		/// <summary>
		/// The Framework will call this frunction from the Addin which every Addin Should Implement
		/// </summary>
		void InitializeAddin(ref IProjectFrameworkApp refProjectFrameworkApp, long lSession,bool bFirstLoad);
		void UnInitializeAddin(long lSession);
	}
}
