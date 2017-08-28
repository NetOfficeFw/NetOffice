using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
	/// <summary>
	/// DispatchInterface NewFile 
	/// SupportByVersion Office, 10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862417.aspx </remarks>
	[SupportByVersion("Office", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class NewFile : _IMsoDispObj
	{
		#pragma warning disable

		#region Type Information

		/// <summary>
		/// Instance Type
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
		public override Type InstanceType
		{
			get
			{
				return LateBindingApiWrapperType;
			}
		}

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
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public NewFile(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

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
		
		/// <param name="progId">registered progID</param>
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
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860279.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="section">optional object section</param>
		/// <param name="displayName">optional object displayName</param>
		/// <param name="action">optional object action</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public bool Add(string fileName, object section, object displayName, object action)
		{
			return Factory.ExecuteBoolMethodGet(this, "Add", fileName, section, displayName, action);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860279.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[CustomMethod]
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public bool Add(string fileName)
		{
			return Factory.ExecuteBoolMethodGet(this, "Add", fileName);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860279.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="section">optional object section</param>
		[CustomMethod]
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public bool Add(string fileName, object section)
		{
			return Factory.ExecuteBoolMethodGet(this, "Add", fileName, section);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860279.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="section">optional object section</param>
		/// <param name="displayName">optional object displayName</param>
		[CustomMethod]
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public bool Add(string fileName, object section, object displayName)
		{
			return Factory.ExecuteBoolMethodGet(this, "Add", fileName, section, displayName);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860573.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="section">optional object section</param>
		/// <param name="displayName">optional object displayName</param>
		/// <param name="action">optional object action</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public bool Remove(string fileName, object section, object displayName, object action)
		{
			return Factory.ExecuteBoolMethodGet(this, "Remove", fileName, section, displayName, action);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860573.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[CustomMethod]
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public bool Remove(string fileName)
		{
			return Factory.ExecuteBoolMethodGet(this, "Remove", fileName);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860573.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="section">optional object section</param>
		[CustomMethod]
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public bool Remove(string fileName, object section)
		{
			return Factory.ExecuteBoolMethodGet(this, "Remove", fileName, section);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860573.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="section">optional object section</param>
		/// <param name="displayName">optional object displayName</param>
		[CustomMethod]
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public bool Remove(string fileName, object section, object displayName)
		{
			return Factory.ExecuteBoolMethodGet(this, "Remove", fileName, section, displayName);
		}

		#endregion

		#pragma warning restore
	}
}
