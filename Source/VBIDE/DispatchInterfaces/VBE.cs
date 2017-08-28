using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VBIDEApi
{
	/// <summary>
	/// DispatchInterface VBE 
	/// SupportByVersion VBIDE, 12,14,5.3
	/// </summary>
	[SupportByVersion("VBIDE", 12,14,5.3)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class VBE : Application
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
                    _type = typeof(VBE);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public VBE(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public VBE(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public VBE(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public VBE(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public VBE(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public VBE(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public VBE() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public VBE(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public NetOffice.VBIDEApi.VBProjects VBProjects
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.VBProjects>(this, "VBProjects", NetOffice.VBIDEApi.VBProjects.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public NetOffice.OfficeApi.CommandBars CommandBars
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CommandBars>(this, "CommandBars", NetOffice.OfficeApi.CommandBars.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public NetOffice.VBIDEApi.CodePanes CodePanes
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.CodePanes>(this, "CodePanes", NetOffice.VBIDEApi.CodePanes.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public NetOffice.VBIDEApi.Windows Windows
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.Windows>(this, "Windows", NetOffice.VBIDEApi.Windows.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public NetOffice.VBIDEApi.Events Events
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.Events>(this, "Events", NetOffice.VBIDEApi.Events.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get/Set
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public NetOffice.VBIDEApi.VBProject ActiveVBProject
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.VBProject>(this, "ActiveVBProject", NetOffice.VBIDEApi.VBProject.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "ActiveVBProject", value);
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public NetOffice.VBIDEApi.VBComponent SelectedVBComponent
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.VBComponent>(this, "SelectedVBComponent", NetOffice.VBIDEApi.VBComponent.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public NetOffice.VBIDEApi.Window MainWindow
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.Window>(this, "MainWindow", NetOffice.VBIDEApi.Window.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public NetOffice.VBIDEApi.Window ActiveWindow
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.Window>(this, "ActiveWindow", NetOffice.VBIDEApi.Window.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get/Set
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public NetOffice.VBIDEApi.CodePane ActiveCodePane
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.CodePane>(this, "ActiveCodePane", NetOffice.VBIDEApi.CodePane.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "ActiveCodePane", value);
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public NetOffice.VBIDEApi.Addins Addins
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.Addins>(this, "Addins", NetOffice.VBIDEApi.Addins.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}
