using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface INavigationControl 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class INavigationControl : COMObject
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
                    _type = typeof(INavigationControl);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public INavigationControl(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public INavigationControl(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public INavigationControl(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public INavigationControl(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public INavigationControl(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public INavigationControl(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public INavigationControl() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public INavigationControl(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.MSDATASRCApi.DataSource DataSource
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSDATASRCApi.DataSource>(this, "DataSource", NetOffice.MSDATASRCApi.DataSource.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "DataSource", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string RecordSource
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "RecordSource");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RecordSource", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string RecordsetLabel
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "RecordsetLabel");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RecordsetLabel", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool ShowFirstButton
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowFirstButton");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowFirstButton", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool ShowPrevButton
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowPrevButton");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowPrevButton", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool ShowNextButton
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowNextButton");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowNextButton", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool ShowLastButton
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowLastButton");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowLastButton", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool ShowNewButton
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowNewButton");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowNewButton", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool ShowDelButton
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowDelButton");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowDelButton", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool ShowSaveButton
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowSaveButton");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowSaveButton", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool ShowUndoButton
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowUndoButton");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowUndoButton", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool ShowSortAscendingButton
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowSortAscendingButton");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowSortAscendingButton", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool ShowSortDescendingButton
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowSortDescendingButton");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowSortDescendingButton", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool ShowFilterBySelectionButton
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowFilterBySelectionButton");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowFilterBySelectionButton", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool ShowToggleFilterButton
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowToggleFilterButton");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowToggleFilterButton", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool ShowHelpButton
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowHelpButton");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowHelpButton", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool ShowLabel
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowLabel");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowLabel", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string FontName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FontName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FontName", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string _State
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "_State");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "_State", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="navbtn">NetOffice.OWC10Api.Enums.NavButtonEnum navbtn</param>
		[SupportByVersion("OWC10", 1)]
		public bool IsButtonEnabled(NetOffice.OWC10Api.Enums.NavButtonEnum navbtn)
		{
			return Factory.ExecuteBoolMethodGet(this, "IsButtonEnabled", navbtn);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public void Redraw()
		{
			 Factory.ExecuteMethod(this, "Redraw");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public void ReleaseDataPage()
		{
			 Factory.ExecuteMethod(this, "ReleaseDataPage");
		}

		#endregion

		#pragma warning restore
	}
}
