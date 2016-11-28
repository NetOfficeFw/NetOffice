using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.OfficeApi
{
	///<summary>
	/// DispatchInterface NewFile 
	/// SupportByVersion Office, 10,11,12,14,15,16
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862417.aspx
	///</summary>
	[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class NewFile : _IMsoDispObj
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
                    _type = typeof(NewFile);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public NewFile(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NewFile(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NewFile(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NewFile(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NewFile(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NewFile() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NewFile(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860279.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="section">optional object Section</param>
		/// <param name="displayName">optional object DisplayName</param>
		/// <param name="action">optional object Action</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public bool Add(string fileName, object section, object displayName, object action)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, section, displayName, action);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860279.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public bool Add(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860279.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="section">optional object Section</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public bool Add(string fileName, object section)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, section);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860279.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="section">optional object Section</param>
		/// <param name="displayName">optional object DisplayName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public bool Add(string fileName, object section, object displayName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, section, displayName);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860573.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="section">optional object Section</param>
		/// <param name="displayName">optional object DisplayName</param>
		/// <param name="action">optional object Action</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public bool Remove(string fileName, object section, object displayName, object action)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, section, displayName, action);
			object returnItem = Invoker.MethodReturn(this, "Remove", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860573.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public bool Remove(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			object returnItem = Invoker.MethodReturn(this, "Remove", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860573.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="section">optional object Section</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public bool Remove(string fileName, object section)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, section);
			object returnItem = Invoker.MethodReturn(this, "Remove", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860573.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="section">optional object Section</param>
		/// <param name="displayName">optional object DisplayName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public bool Remove(string fileName, object section, object displayName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, section, displayName);
			object returnItem = Invoker.MethodReturn(this, "Remove", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		#endregion
		#pragma warning restore
	}
}