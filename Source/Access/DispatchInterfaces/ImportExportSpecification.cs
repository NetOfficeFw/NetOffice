using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// DispatchInterface ImportExportSpecification 
	/// SupportByVersion Access, 12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.ImportExportSpecification"/> </remarks>
	[SupportByVersion("Access", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class ImportExportSpecification : COMObject
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
                    _type = typeof(ImportExportSpecification);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public ImportExportSpecification(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ImportExportSpecification(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ImportExportSpecification(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ImportExportSpecification(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ImportExportSpecification(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ImportExportSpecification(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ImportExportSpecification() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ImportExportSpecification(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.ImportExportSpecification.Application"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		public NetOffice.AccessApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Application>(this, "Application", NetOffice.AccessApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.ImportExportSpecification.Parent"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.ImportExportSpecification.Name"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Name", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.ImportExportSpecification.XML"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		public string XML
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "XML");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "XML", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.ImportExportSpecification.Description"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		public string Description
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Description");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Description", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.ImportExportSpecification.Execute"/> </remarks>
		/// <param name="prompt">optional object prompt</param>
		[SupportByVersion("Access", 12,14,15,16)]
		public void Execute(object prompt)
		{
			 Factory.ExecuteMethod(this, "Execute", prompt);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.ImportExportSpecification.Execute"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public void Execute()
		{
			 Factory.ExecuteMethod(this, "Execute");
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.ImportExportSpecification.Delete"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="dispid">Int32 dispid</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 12,14,15,16)]
		public bool IsMemberSafe(Int32 dispid)
		{
			return Factory.ExecuteBoolMethodGet(this, "IsMemberSafe", dispid);
		}

		#endregion

		#pragma warning restore
	}
}
