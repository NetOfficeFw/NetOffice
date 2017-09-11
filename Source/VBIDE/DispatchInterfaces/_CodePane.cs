using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VBIDEApi
{
	/// <summary>
	/// DispatchInterface _CodePane 
	/// SupportByVersion VBIDE, 12,14,5.3
	/// </summary>
	[SupportByVersion("VBIDE", 12,14,5.3)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _CodePane : COMObject
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
                    _type = typeof(_CodePane);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _CodePane(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _CodePane(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CodePane(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CodePane(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CodePane(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CodePane(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CodePane() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CodePane(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public NetOffice.VBIDEApi.CodePanes Collection
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.CodePanes>(this, "Collection", NetOffice.VBIDEApi.CodePanes.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public NetOffice.VBIDEApi.VBE VBE
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.VBE>(this, "VBE", NetOffice.VBIDEApi.VBE.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public NetOffice.VBIDEApi.Window Window
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.Window>(this, "Window", NetOffice.VBIDEApi.Window.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get/Set
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public Int32 TopLine
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "TopLine");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TopLine", value);
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public Int32 CountOfVisibleLines
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "CountOfVisibleLines");
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public NetOffice.VBIDEApi.CodeModule CodeModule
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.CodeModule>(this, "CodeModule", NetOffice.VBIDEApi.CodeModule.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public NetOffice.VBIDEApi.Enums.vbext_CodePaneview CodePaneView
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.VBIDEApi.Enums.vbext_CodePaneview>(this, "CodePaneView");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// </summary>
		/// <param name="startLine">Int32 startLine</param>
		/// <param name="startColumn">Int32 startColumn</param>
		/// <param name="endLine">Int32 endLine</param>
		/// <param name="endColumn">Int32 endColumn</param>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public void GetSelection(out Int32 startLine, out Int32 startColumn, out Int32 endLine, out Int32 endColumn)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true,true,true,true);
			startLine = 0;
			startColumn = 0;
			endLine = 0;
			endColumn = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(startLine, startColumn, endLine, endColumn);
			Invoker.Method(this, "GetSelection", paramsArray, modifiers);
			startLine = (Int32)paramsArray[0];
			startColumn = (Int32)paramsArray[1];
			endLine = (Int32)paramsArray[2];
			endColumn = (Int32)paramsArray[3];
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// </summary>
		/// <param name="startLine">Int32 startLine</param>
		/// <param name="startColumn">Int32 startColumn</param>
		/// <param name="endLine">Int32 endLine</param>
		/// <param name="endColumn">Int32 endColumn</param>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public void SetSelection(Int32 startLine, Int32 startColumn, Int32 endLine, Int32 endColumn)
		{
			 Factory.ExecuteMethod(this, "SetSelection", startLine, startColumn, endLine, endColumn);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public void Show()
		{
			 Factory.ExecuteMethod(this, "Show");
		}

		#endregion

		#pragma warning restore
	}
}
