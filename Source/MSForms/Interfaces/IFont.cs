using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSFormsApi
{
	/// <summary>
	/// Interface IFont 
	/// SupportByVersion MSForms, 2
	/// </summary>
	[SupportByVersion("MSForms", 2)]
	[EntityType(EntityType.IsInterface)]
 	public class IFont : COMObject
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
                    _type = typeof(IFont);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IFont(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IFont(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IFont(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IFont(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IFont(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IFont(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IFont() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IFont(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
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
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public float Size
		{
			get
			{
				return Factory.ExecuteFloatPropertyGet(this, "Size");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Size", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public bool Bold
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Bold");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Bold", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public bool Italic
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Italic");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Italic", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public bool Underline
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Underline");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Underline", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public bool Strikethrough
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Strikethrough");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Strikethrough", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public Int16 Weight
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Weight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Weight", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public Int16 Charset
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Charset");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Charset", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public Int32 hFont
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "hFont");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="lplpfont">NetOffice.MSFormsApi.IFont lplpfont</param>
		[SupportByVersion("MSForms", 2)]
		public Int32 Clone(out NetOffice.MSFormsApi.IFont lplpfont)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			lplpfont = null;
			object[] paramsArray = Invoker.ValidateParamsArray(new object());
			object returnItem = Invoker.MethodReturn(this, "Clone", paramsArray, modifiers);
            if (paramsArray[1] is MarshalByRefObject)
                paramsArray[1] = new NetOffice.MSFormsApi.IFont(this, paramsArray[1]);
            else
                paramsArray[1] = null;
            lplpfont = (NetOffice.MSFormsApi.IFont)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="lpFontOther">NetOffice.MSFormsApi.IFont lpFontOther</param>
		[SupportByVersion("MSForms", 2)]
		public Int32 IsEqual(NetOffice.MSFormsApi.IFont lpFontOther)
		{
			return Factory.ExecuteInt32MethodGet(this, "IsEqual", lpFontOther);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="cyLogical">Int32 cyLogical</param>
		/// <param name="cyHimetric">Int32 cyHimetric</param>
		[SupportByVersion("MSForms", 2)]
		public Int32 SetRatio(Int32 cyLogical, Int32 cyHimetric)
		{
			return Factory.ExecuteInt32MethodGet(this, "SetRatio", cyLogical, cyHimetric);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="hFont">Int32 hFont</param>
		[SupportByVersion("MSForms", 2)]
		public Int32 AddRefHfont(Int32 hFont)
		{
			return Factory.ExecuteInt32MethodGet(this, "AddRefHfont", hFont);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="hFont">Int32 hFont</param>
		[SupportByVersion("MSForms", 2)]
		public Int32 ReleaseHfont(Int32 hFont)
		{
			return Factory.ExecuteInt32MethodGet(this, "ReleaseHfont", hFont);
		}

		#endregion

		#pragma warning restore
	}
}
