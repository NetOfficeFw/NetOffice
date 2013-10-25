using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using NetOffice;
namespace NetOffice.WordApi
{
	///<summary>
	/// DispatchInterface ProtectedViewWindows 
	/// SupportByVersion Word, 14,15
	///</summary>
	///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197163.aspx </remarks>
	[SupportByVersionAttribute("Word", 14,15)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class ProtectedViewWindows : COMObject ,IEnumerable<NetOffice.WordApi.ProtectedViewWindow>
	{
		#pragma warning disable
		#region Type Information

        private static Type _type;

		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(ProtectedViewWindows);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ProtectedViewWindows(Core factory, COMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ProtectedViewWindows(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ProtectedViewWindows(Core factory, COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ProtectedViewWindows(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ProtectedViewWindows(COMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ProtectedViewWindows() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ProtectedViewWindows(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 14, 15
		/// Get
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196633.aspx </remarks>
		[SupportByVersionAttribute("Word", 14,15)]
		public NetOffice.WordApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.WordApi.Application newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Application.LateBindingApiWrapperType) as NetOffice.WordApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15
		/// Get
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821613.aspx </remarks>
		[SupportByVersionAttribute("Word", 14,15)]
		public Int32 Creator
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Creator", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822106.aspx </remarks>
		[SupportByVersionAttribute("Word", 14,15)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				COMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15
		/// Get
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840128.aspx </remarks>
		[SupportByVersionAttribute("Word", 14,15)]
		public Int32 Count
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Count", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 14, 15
		/// </summary>
		/// <param name="index">object Index</param>
		[SupportByVersionAttribute("Word", 14,15)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.WordApi.ProtectedViewWindow this[object index]
		{
			get
			{
				object[] paramsArray = Invoker.ValidateParamsArray(index);
				object returnItem = Invoker.MethodReturn(this, "Item", paramsArray);
				NetOffice.WordApi.ProtectedViewWindow newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.ProtectedViewWindow.LateBindingApiWrapperType) as NetOffice.WordApi.ProtectedViewWindow;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="visible">optional object Visible</param>
		/// <param name="openAndRepair">optional object OpenAndRepair</param>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193715.aspx </remarks>
		[SupportByVersionAttribute("Word", 14,15)]
		public NetOffice.WordApi.ProtectedViewWindow Open(object fileName, object addToRecentFiles, object passwordDocument, object visible, object openAndRepair)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, addToRecentFiles, passwordDocument, visible, openAndRepair);
			object returnItem = Invoker.MethodReturn(this, "Open", paramsArray);
			NetOffice.WordApi.ProtectedViewWindow newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.ProtectedViewWindow.LateBindingApiWrapperType) as NetOffice.WordApi.ProtectedViewWindow;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 14, 15
		/// </summary>
		/// <param name="fileName">object FileName</param>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193715.aspx </remarks>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15)]
		public NetOffice.WordApi.ProtectedViewWindow Open(object fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			object returnItem = Invoker.MethodReturn(this, "Open", paramsArray);
			NetOffice.WordApi.ProtectedViewWindow newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.ProtectedViewWindow.LateBindingApiWrapperType) as NetOffice.WordApi.ProtectedViewWindow;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 14, 15
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193715.aspx </remarks>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15)]
		public NetOffice.WordApi.ProtectedViewWindow Open(object fileName, object addToRecentFiles)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, addToRecentFiles);
			object returnItem = Invoker.MethodReturn(this, "Open", paramsArray);
			NetOffice.WordApi.ProtectedViewWindow newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.ProtectedViewWindow.LateBindingApiWrapperType) as NetOffice.WordApi.ProtectedViewWindow;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 14, 15
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193715.aspx </remarks>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15)]
		public NetOffice.WordApi.ProtectedViewWindow Open(object fileName, object addToRecentFiles, object passwordDocument)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, addToRecentFiles, passwordDocument);
			object returnItem = Invoker.MethodReturn(this, "Open", paramsArray);
			NetOffice.WordApi.ProtectedViewWindow newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.ProtectedViewWindow.LateBindingApiWrapperType) as NetOffice.WordApi.ProtectedViewWindow;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 14, 15
		/// </summary>
		/// <param name="fileName">object FileName</param>
		/// <param name="addToRecentFiles">optional object AddToRecentFiles</param>
		/// <param name="passwordDocument">optional object PasswordDocument</param>
		/// <param name="visible">optional object Visible</param>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193715.aspx </remarks>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15)]
		public NetOffice.WordApi.ProtectedViewWindow Open(object fileName, object addToRecentFiles, object passwordDocument, object visible)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, addToRecentFiles, passwordDocument, visible);
			object returnItem = Invoker.MethodReturn(this, "Open", paramsArray);
			NetOffice.WordApi.ProtectedViewWindow newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.ProtectedViewWindow.LateBindingApiWrapperType) as NetOffice.WordApi.ProtectedViewWindow;
			return newObject;
		}

		#endregion

       #region IEnumerable<NetOffice.WordApi.ProtectedViewWindow> Member
        
        /// <summary>
		/// SupportByVersionAttribute Word, 14,15
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15)]
       public IEnumerator<NetOffice.WordApi.ProtectedViewWindow> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.WordApi.ProtectedViewWindow item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersionAttribute Word, 14,15
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion
		#pragma warning restore
	}
}