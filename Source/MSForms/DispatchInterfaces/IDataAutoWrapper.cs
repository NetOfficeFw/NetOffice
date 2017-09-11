using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSFormsApi
{
	/// <summary>
	/// DispatchInterface IDataAutoWrapper 
	/// SupportByVersion MSForms, 2
	/// </summary>
	[SupportByVersion("MSForms", 2)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IDataAutoWrapper : COMObject
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
                    _type = typeof(IDataAutoWrapper);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IDataAutoWrapper(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IDataAutoWrapper(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDataAutoWrapper(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDataAutoWrapper(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDataAutoWrapper(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDataAutoWrapper(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDataAutoWrapper() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDataAutoWrapper(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public void Clear()
		{
			 Factory.ExecuteMethod(this, "Clear");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="format">object format</param>
		[SupportByVersion("MSForms", 2)]
		public bool GetFormat(object format)
		{
			return Factory.ExecuteBoolMethodGet(this, "GetFormat", format);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="format">optional object format</param>
		[SupportByVersion("MSForms", 2)]
		public string GetText(object format)
		{
			return Factory.ExecuteStringMethodGet(this, "GetText", format);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		public string GetText()
		{
			return Factory.ExecuteStringMethodGet(this, "GetText");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="text">string text</param>
		/// <param name="format">optional object format</param>
		[SupportByVersion("MSForms", 2)]
		public void SetText(string text, object format)
		{
			 Factory.ExecuteMethod(this, "SetText", text, format);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="text">string text</param>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		public void SetText(string text)
		{
			 Factory.ExecuteMethod(this, "SetText", text);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public void PutInClipboard()
		{
			 Factory.ExecuteMethod(this, "PutInClipboard");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public void GetFromClipboard()
		{
			 Factory.ExecuteMethod(this, "GetFromClipboard");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="oKEffect">optional object oKEffect</param>
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Enums.fmDropEffect StartDrag(object oKEffect)
		{
			return Factory.ExecuteEnumMethodGet<NetOffice.MSFormsApi.Enums.fmDropEffect>(this, "StartDrag", oKEffect);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Enums.fmDropEffect StartDrag()
		{
			return Factory.ExecuteEnumMethodGet<NetOffice.MSFormsApi.Enums.fmDropEffect>(this, "StartDrag");
		}

		#endregion

		#pragma warning restore
	}
}
