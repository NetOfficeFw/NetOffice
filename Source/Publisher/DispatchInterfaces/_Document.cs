using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.PublisherApi
{
	///<summary>
	/// DispatchInterface _Document 
	/// SupportByVersion Publisher, 14,15,16
	///</summary>
	[SupportByVersionAttribute("Publisher", 14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class _Document : COMObject
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
                    _type = typeof(_Document);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _Document(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Document(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Document(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Document(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Document(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Document() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Document(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string ActivePrinter
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ActivePrinter", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ActivePrinter", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Window ActiveWindow
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ActiveWindow", paramsArray);
				NetOffice.PublisherApi.Window newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.Window.LateBindingApiWrapperType) as NetOffice.PublisherApi.Window;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.PublisherApi.Application newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.Application.LateBindingApiWrapperType) as NetOffice.PublisherApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Enums.PbColorMode ColorMode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ColorMode", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.PublisherApi.Enums.PbColorMode)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.ColorScheme ColorScheme
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ColorScheme", paramsArray);
				NetOffice.PublisherApi.ColorScheme newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.ColorScheme.LateBindingApiWrapperType) as NetOffice.PublisherApi.ColorScheme;
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ColorScheme", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public object DefaultTabStop
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DefaultTabStop", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DefaultTabStop", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public bool EnvelopeVisible
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EnvelopeVisible", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "EnvelopeVisible", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public string FullName
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FullName", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.LayoutGuides LayoutGuides
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LayoutGuides", paramsArray);
				NetOffice.PublisherApi.LayoutGuides newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.LayoutGuides.LateBindingApiWrapperType) as NetOffice.PublisherApi.LayoutGuides;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.MsoEnvelope MailEnvelope
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MailEnvelope", paramsArray);
				NetOffice.OfficeApi.MsoEnvelope newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.MsoEnvelope.LateBindingApiWrapperType) as NetOffice.OfficeApi.MsoEnvelope;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.MailMerge MailMerge
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MailMerge", paramsArray);
				NetOffice.PublisherApi.MailMerge newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.MailMerge.LateBindingApiWrapperType) as NetOffice.PublisherApi.MailMerge;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.MasterPages MasterPages
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MasterPages", paramsArray);
				NetOffice.PublisherApi.MasterPages newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.MasterPages.LateBindingApiWrapperType) as NetOffice.PublisherApi.MasterPages;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public string Name
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Name", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Pages Pages
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Pages", paramsArray);
				NetOffice.PublisherApi.Pages newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.Pages.LateBindingApiWrapperType) as NetOffice.PublisherApi.Pages;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.PageSetup PageSetup
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PageSetup", paramsArray);
				NetOffice.PublisherApi.PageSetup newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.PageSetup.LateBindingApiWrapperType) as NetOffice.PublisherApi.PageSetup;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public string Path
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Path", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.PublisherApi.Enums.PbPersonalInfoSet PersonalInformationSet
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PersonalInformationSet", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.PublisherApi.Enums.PbPersonalInfoSet)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "PersonalInformationSet", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Plates Plates
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Plates", paramsArray);
				NetOffice.PublisherApi.Plates newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.Plates.LateBindingApiWrapperType) as NetOffice.PublisherApi.Plates;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public bool ReadOnly
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ReadOnly", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Enums.PbDirectionType DocumentDirection
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DocumentDirection", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.PublisherApi.Enums.PbDirectionType)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DocumentDirection", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public bool Saved
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Saved", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Enums.PbFileFormat SaveFormat
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SaveFormat", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.PublisherApi.Enums.PbFileFormat)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.ScratchArea ScratchArea
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ScratchArea", paramsArray);
				NetOffice.PublisherApi.ScratchArea newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.ScratchArea.LateBindingApiWrapperType) as NetOffice.PublisherApi.ScratchArea;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Selection Selection
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Selection", paramsArray);
				NetOffice.PublisherApi.Selection newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.Selection.LateBindingApiWrapperType) as NetOffice.PublisherApi.Selection;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Stories Stories
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Stories", paramsArray);
				NetOffice.PublisherApi.Stories newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.Stories.LateBindingApiWrapperType) as NetOffice.PublisherApi.Stories;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Tags Tags
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Tags", paramsArray);
				NetOffice.PublisherApi.Tags newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.Tags.LateBindingApiWrapperType) as NetOffice.PublisherApi.Tags;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextStyles TextStyles
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TextStyles", paramsArray);
				NetOffice.PublisherApi.TextStyles newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.TextStyles.LateBindingApiWrapperType) as NetOffice.PublisherApi.TextStyles;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool ViewBoundariesAndGuides
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ViewBoundariesAndGuides", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ViewBoundariesAndGuides", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public bool ViewTwoPageSpread
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ViewTwoPageSpread", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ViewTwoPageSpread", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Wizard Wizard
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Wizard", paramsArray);
				NetOffice.PublisherApi.Wizard newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.Wizard.LateBindingApiWrapperType) as NetOffice.PublisherApi.Wizard;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.View ActiveView
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ActiveView", paramsArray);
				NetOffice.PublisherApi.View newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.View.LateBindingApiWrapperType) as NetOffice.PublisherApi.View;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.AdvancedPrintOptions AdvancedPrintOptions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AdvancedPrintOptions", paramsArray);
				NetOffice.PublisherApi.AdvancedPrintOptions newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.AdvancedPrintOptions.LateBindingApiWrapperType) as NetOffice.PublisherApi.AdvancedPrintOptions;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.BorderArts BorderArts
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BorderArts", paramsArray);
				NetOffice.PublisherApi.BorderArts newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.BorderArts.LateBindingApiWrapperType) as NetOffice.PublisherApi.BorderArts;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public bool IsDataSourceConnected
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IsDataSourceConnected", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.FindReplace Find
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Find", paramsArray);
				NetOffice.PublisherApi.FindReplace newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.FindReplace.LateBindingApiWrapperType) as NetOffice.PublisherApi.FindReplace;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public Int32 UndoActionsAvailable
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "UndoActionsAvailable", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public Int32 RedoActionsAvailable
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RedoActionsAvailable", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public bool ViewHorizontalBaseLineGuides
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ViewHorizontalBaseLineGuides", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ViewHorizontalBaseLineGuides", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public bool ViewVerticalBaseLineGuides
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ViewVerticalBaseLineGuides", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ViewVerticalBaseLineGuides", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Enums.PbPublicationType PublicationType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PublicationType", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.PublisherApi.Enums.PbPublicationType)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Sections Sections
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Sections", paramsArray);
				NetOffice.PublisherApi.Sections newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.Sections.LateBindingApiWrapperType) as NetOffice.PublisherApi.Sections;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.WebNavigationBarSets WebNavigationBarSets
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "WebNavigationBarSets", paramsArray);
				NetOffice.PublisherApi.WebNavigationBarSets newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.WebNavigationBarSets.LateBindingApiWrapperType) as NetOffice.PublisherApi.WebNavigationBarSets;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public bool RemovePersonalInformation
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RemovePersonalInformation", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "RemovePersonalInformation", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public bool PrintPageBackgrounds
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PrintPageBackgrounds", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "PrintPageBackgrounds", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.ColorsInUse ColorsInUse
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ColorsInUse", paramsArray);
				NetOffice.PublisherApi.ColorsInUse newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.ColorsInUse.LateBindingApiWrapperType) as NetOffice.PublisherApi.ColorsInUse;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public bool IsWizard
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IsWizard", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.ShapeRange SurplusShapes
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SurplusShapes", paramsArray);
				NetOffice.PublisherApi.ShapeRange newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.ShapeRange.LateBindingApiWrapperType) as NetOffice.PublisherApi.ShapeRange;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Enums.PbPrintStyle PrintStyle
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PrintStyle", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.PublisherApi.Enums.PbPrintStyle)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public bool ViewBoundaries
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ViewBoundaries", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ViewBoundaries", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public bool ViewGuides
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ViewGuides", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ViewGuides", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.BuildingBlocks AvailableBuildingBlocks
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AvailableBuildingBlocks", paramsArray);
				NetOffice.PublisherApi.BuildingBlocks newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.BuildingBlocks.LateBindingApiWrapperType) as NetOffice.PublisherApi.BuildingBlocks;
				return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void Close()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Close", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="mode">NetOffice.PublisherApi.Enums.PbColorMode Mode</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Plates CreatePlateCollection(NetOffice.PublisherApi.Enums.PbColorMode mode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(mode);
			object returnItem = Invoker.MethodReturn(this, "CreatePlateCollection", paramsArray);
			NetOffice.PublisherApi.Plates newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Plates.LateBindingApiWrapperType) as NetOffice.PublisherApi.Plates;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="mode">NetOffice.PublisherApi.Enums.PbColorMode Mode</param>
		/// <param name="plates">optional object Plates</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void EnterColorMode10(NetOffice.PublisherApi.Enums.PbColorMode mode, object plates)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(mode, plates);
			Invoker.Method(this, "EnterColorMode10", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="mode">NetOffice.PublisherApi.Enums.PbColorMode Mode</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void EnterColorMode10(NetOffice.PublisherApi.Enums.PbColorMode mode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(mode);
			Invoker.Method(this, "EnterColorMode10", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="tagName">string TagName</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.ShapeRange FindShapesByTag(string tagName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(tagName);
			object returnItem = Invoker.MethodReturn(this, "FindShapesByTag", paramsArray);
			NetOffice.PublisherApi.ShapeRange newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.ShapeRange.LateBindingApiWrapperType) as NetOffice.PublisherApi.ShapeRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="wizardTag">NetOffice.PublisherApi.Enums.PbWizardTag WizardTag</param>
		/// <param name="instance">optional Int32 Instance = -1</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.ShapeRange FindShapeByWizardTag(NetOffice.PublisherApi.Enums.PbWizardTag wizardTag, object instance)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wizardTag, instance);
			object returnItem = Invoker.MethodReturn(this, "FindShapeByWizardTag", paramsArray);
			NetOffice.PublisherApi.ShapeRange newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.ShapeRange.LateBindingApiWrapperType) as NetOffice.PublisherApi.ShapeRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="wizardTag">NetOffice.PublisherApi.Enums.PbWizardTag WizardTag</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.ShapeRange FindShapeByWizardTag(NetOffice.PublisherApi.Enums.PbWizardTag wizardTag)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wizardTag);
			object returnItem = Invoker.MethodReturn(this, "FindShapeByWizardTag", paramsArray);
			NetOffice.PublisherApi.ShapeRange newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.ShapeRange.LateBindingApiWrapperType) as NetOffice.PublisherApi.ShapeRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="printToFile">optional string PrintToFile = </param>
		/// <param name="copies">optional Int32 Copies = -1</param>
		/// <param name="collate">optional bool Collate = true</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void PrintOut(object from, object to, object printToFile, object copies, object collate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, printToFile, copies, collate);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void PrintOut()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void PrintOut(object from)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void PrintOut(object from, object to)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="printToFile">optional string PrintToFile = </param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void PrintOut(object from, object to, object printToFile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, printToFile);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="printToFile">optional string PrintToFile = </param>
		/// <param name="copies">optional Int32 Copies = -1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void PrintOut(object from, object to, object printToFile, object copies)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, printToFile, copies);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void Save()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Save", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">optional object Filename</param>
		/// <param name="format">optional NetOffice.PublisherApi.Enums.PbFileFormat Format = 1</param>
		/// <param name="addToRecentFiles">optional bool AddToRecentFiles = true</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void SaveAs(object filename, object format, object addToRecentFiles)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, format, addToRecentFiles);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void SaveAs()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">optional object Filename</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void SaveAs(object filename)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">optional object Filename</param>
		/// <param name="format">optional NetOffice.PublisherApi.Enums.PbFileFormat Format = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void SaveAs(object filename, object format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, format);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="oh">Int32 oh</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void SelectID(Int32 oh)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(oh);
			Invoker.Method(this, "SelectID", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void UndoClear()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "UndoClear", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void UpdateOLEObjects()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "UpdateOLEObjects", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="count">optional Int32 Count = 1</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void Undo(object count)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(count);
			Invoker.Method(this, "Undo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void Undo()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Undo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="count">optional Int32 Count = 1</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void Redo(object count)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(count);
			Invoker.Method(this, "Redo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void Redo()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Redo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="actionName">string ActionName</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void BeginCustomUndoAction(string actionName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(actionName);
			Invoker.Method(this, "BeginCustomUndoAction", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void EndCustomUndoAction()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "EndCustomUndoAction", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void WebPagePreview()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "WebPagePreview", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="value">NetOffice.PublisherApi.Enums.PbPublicationType Value</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ConvertPublicationType(NetOffice.PublisherApi.Enums.PbPublicationType value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(value);
			Invoker.Method(this, "ConvertPublicationType", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="mode">NetOffice.PublisherApi.Enums.PbColorMode Mode</param>
		/// <param name="plates">optional object Plates</param>
		/// <param name="deleteExcessInks">optional bool DeleteExcessInks = false</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void EnterColorMode(NetOffice.PublisherApi.Enums.PbColorMode mode, object plates, object deleteExcessInks)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(mode, plates, deleteExcessInks);
			Invoker.Method(this, "EnterColorMode", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="mode">NetOffice.PublisherApi.Enums.PbColorMode Mode</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void EnterColorMode(NetOffice.PublisherApi.Enums.PbColorMode mode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(mode);
			Invoker.Method(this, "EnterColorMode", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="mode">NetOffice.PublisherApi.Enums.PbColorMode Mode</param>
		/// <param name="plates">optional object Plates</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void EnterColorMode(NetOffice.PublisherApi.Enums.PbColorMode mode, object plates)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(mode, plates);
			Invoker.Method(this, "EnterColorMode", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="printToFile">optional string PrintToFile = </param>
		/// <param name="copies">optional Int32 Copies = -1</param>
		/// <param name="collate">optional bool Collate = true</param>
		/// <param name="printStyle">optional NetOffice.PublisherApi.Enums.PbPrintStyle PrintStyle = 0</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void PrintOutEx(object from, object to, object printToFile, object copies, object collate, object printStyle)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, printToFile, copies, collate, printStyle);
			Invoker.Method(this, "PrintOutEx", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void PrintOutEx()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "PrintOutEx", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void PrintOutEx(object from)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from);
			Invoker.Method(this, "PrintOutEx", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void PrintOutEx(object from, object to)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to);
			Invoker.Method(this, "PrintOutEx", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="printToFile">optional string PrintToFile = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void PrintOutEx(object from, object to, object printToFile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, printToFile);
			Invoker.Method(this, "PrintOutEx", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="printToFile">optional string PrintToFile = </param>
		/// <param name="copies">optional Int32 Copies = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void PrintOutEx(object from, object to, object printToFile, object copies)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, printToFile, copies);
			Invoker.Method(this, "PrintOutEx", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="printToFile">optional string PrintToFile = </param>
		/// <param name="copies">optional Int32 Copies = -1</param>
		/// <param name="collate">optional bool Collate = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void PrintOutEx(object from, object to, object printToFile, object copies, object collate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, printToFile, copies, collate);
			Invoker.Method(this, "PrintOutEx", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="wizard">NetOffice.PublisherApi.Enums.PbWizard Wizard</param>
		/// <param name="design">optional Int32 Design = -1</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ChangeDocument(NetOffice.PublisherApi.Enums.PbWizard wizard, object design)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wizard, design);
			Invoker.Method(this, "ChangeDocument", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="wizard">NetOffice.PublisherApi.Enums.PbWizard Wizard</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ChangeDocument(NetOffice.PublisherApi.Enums.PbWizard wizard)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wizard);
			Invoker.Method(this, "ChangeDocument", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void SetBusinessInformation(string name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			Invoker.Method(this, "SetBusinessInformation", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType Format</param>
		/// <param name="filename">string Filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="colorDownsampleTarget">optional Int32 ColorDownsampleTarget = -1</param>
		/// <param name="colorDownsampleThreshold">optional Int32 ColorDownsampleThreshold = -1</param>
		/// <param name="oneBitDownsampleTarget">optional Int32 OneBitDownsampleTarget = -1</param>
		/// <param name="oneBitDownsampleThreshold">optional Int32 OneBitDownsampleThreshold = -1</param>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="copies">optional Int32 Copies = -1</param>
		/// <param name="collate">optional bool Collate = true</param>
		/// <param name="printStyle">optional NetOffice.PublisherApi.Enums.PbPrintStyle PrintStyle = 0</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		/// <param name="bitmapMissingFonts">optional bool BitmapMissingFonts = true</param>
		/// <param name="useISO19005_1">optional bool UseISO19005_1 = false</param>
		/// <param name="externalExporter">optional object ExternalExporter</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget, object oneBitDownsampleThreshold, object from, object to, object copies, object collate, object printStyle, object docStructureTags, object bitmapMissingFonts, object useISO19005_1, object externalExporter)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, filename, intent, includeDocumentProperties, colorDownsampleTarget, colorDownsampleThreshold, oneBitDownsampleTarget, oneBitDownsampleThreshold, from, to, copies, collate, printStyle, docStructureTags, bitmapMissingFonts, useISO19005_1, externalExporter);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType Format</param>
		/// <param name="filename">string Filename</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, filename);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType Format</param>
		/// <param name="filename">string Filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, filename, intent);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType Format</param>
		/// <param name="filename">string Filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, filename, intent, includeDocumentProperties);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType Format</param>
		/// <param name="filename">string Filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="colorDownsampleTarget">optional Int32 ColorDownsampleTarget = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, filename, intent, includeDocumentProperties, colorDownsampleTarget);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType Format</param>
		/// <param name="filename">string Filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="colorDownsampleTarget">optional Int32 ColorDownsampleTarget = -1</param>
		/// <param name="colorDownsampleThreshold">optional Int32 ColorDownsampleThreshold = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, filename, intent, includeDocumentProperties, colorDownsampleTarget, colorDownsampleThreshold);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType Format</param>
		/// <param name="filename">string Filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="colorDownsampleTarget">optional Int32 ColorDownsampleTarget = -1</param>
		/// <param name="colorDownsampleThreshold">optional Int32 ColorDownsampleThreshold = -1</param>
		/// <param name="oneBitDownsampleTarget">optional Int32 OneBitDownsampleTarget = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, filename, intent, includeDocumentProperties, colorDownsampleTarget, colorDownsampleThreshold, oneBitDownsampleTarget);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType Format</param>
		/// <param name="filename">string Filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="colorDownsampleTarget">optional Int32 ColorDownsampleTarget = -1</param>
		/// <param name="colorDownsampleThreshold">optional Int32 ColorDownsampleThreshold = -1</param>
		/// <param name="oneBitDownsampleTarget">optional Int32 OneBitDownsampleTarget = -1</param>
		/// <param name="oneBitDownsampleThreshold">optional Int32 OneBitDownsampleThreshold = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget, object oneBitDownsampleThreshold)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, filename, intent, includeDocumentProperties, colorDownsampleTarget, colorDownsampleThreshold, oneBitDownsampleTarget, oneBitDownsampleThreshold);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType Format</param>
		/// <param name="filename">string Filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="colorDownsampleTarget">optional Int32 ColorDownsampleTarget = -1</param>
		/// <param name="colorDownsampleThreshold">optional Int32 ColorDownsampleThreshold = -1</param>
		/// <param name="oneBitDownsampleTarget">optional Int32 OneBitDownsampleTarget = -1</param>
		/// <param name="oneBitDownsampleThreshold">optional Int32 OneBitDownsampleThreshold = -1</param>
		/// <param name="from">optional Int32 From = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget, object oneBitDownsampleThreshold, object from)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, filename, intent, includeDocumentProperties, colorDownsampleTarget, colorDownsampleThreshold, oneBitDownsampleTarget, oneBitDownsampleThreshold, from);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType Format</param>
		/// <param name="filename">string Filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="colorDownsampleTarget">optional Int32 ColorDownsampleTarget = -1</param>
		/// <param name="colorDownsampleThreshold">optional Int32 ColorDownsampleThreshold = -1</param>
		/// <param name="oneBitDownsampleTarget">optional Int32 OneBitDownsampleTarget = -1</param>
		/// <param name="oneBitDownsampleThreshold">optional Int32 OneBitDownsampleThreshold = -1</param>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget, object oneBitDownsampleThreshold, object from, object to)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, filename, intent, includeDocumentProperties, colorDownsampleTarget, colorDownsampleThreshold, oneBitDownsampleTarget, oneBitDownsampleThreshold, from, to);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType Format</param>
		/// <param name="filename">string Filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="colorDownsampleTarget">optional Int32 ColorDownsampleTarget = -1</param>
		/// <param name="colorDownsampleThreshold">optional Int32 ColorDownsampleThreshold = -1</param>
		/// <param name="oneBitDownsampleTarget">optional Int32 OneBitDownsampleTarget = -1</param>
		/// <param name="oneBitDownsampleThreshold">optional Int32 OneBitDownsampleThreshold = -1</param>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="copies">optional Int32 Copies = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget, object oneBitDownsampleThreshold, object from, object to, object copies)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, filename, intent, includeDocumentProperties, colorDownsampleTarget, colorDownsampleThreshold, oneBitDownsampleTarget, oneBitDownsampleThreshold, from, to, copies);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType Format</param>
		/// <param name="filename">string Filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="colorDownsampleTarget">optional Int32 ColorDownsampleTarget = -1</param>
		/// <param name="colorDownsampleThreshold">optional Int32 ColorDownsampleThreshold = -1</param>
		/// <param name="oneBitDownsampleTarget">optional Int32 OneBitDownsampleTarget = -1</param>
		/// <param name="oneBitDownsampleThreshold">optional Int32 OneBitDownsampleThreshold = -1</param>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="copies">optional Int32 Copies = -1</param>
		/// <param name="collate">optional bool Collate = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget, object oneBitDownsampleThreshold, object from, object to, object copies, object collate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, filename, intent, includeDocumentProperties, colorDownsampleTarget, colorDownsampleThreshold, oneBitDownsampleTarget, oneBitDownsampleThreshold, from, to, copies, collate);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType Format</param>
		/// <param name="filename">string Filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="colorDownsampleTarget">optional Int32 ColorDownsampleTarget = -1</param>
		/// <param name="colorDownsampleThreshold">optional Int32 ColorDownsampleThreshold = -1</param>
		/// <param name="oneBitDownsampleTarget">optional Int32 OneBitDownsampleTarget = -1</param>
		/// <param name="oneBitDownsampleThreshold">optional Int32 OneBitDownsampleThreshold = -1</param>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="copies">optional Int32 Copies = -1</param>
		/// <param name="collate">optional bool Collate = true</param>
		/// <param name="printStyle">optional NetOffice.PublisherApi.Enums.PbPrintStyle PrintStyle = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget, object oneBitDownsampleThreshold, object from, object to, object copies, object collate, object printStyle)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, filename, intent, includeDocumentProperties, colorDownsampleTarget, colorDownsampleThreshold, oneBitDownsampleTarget, oneBitDownsampleThreshold, from, to, copies, collate, printStyle);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType Format</param>
		/// <param name="filename">string Filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="colorDownsampleTarget">optional Int32 ColorDownsampleTarget = -1</param>
		/// <param name="colorDownsampleThreshold">optional Int32 ColorDownsampleThreshold = -1</param>
		/// <param name="oneBitDownsampleTarget">optional Int32 OneBitDownsampleTarget = -1</param>
		/// <param name="oneBitDownsampleThreshold">optional Int32 OneBitDownsampleThreshold = -1</param>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="copies">optional Int32 Copies = -1</param>
		/// <param name="collate">optional bool Collate = true</param>
		/// <param name="printStyle">optional NetOffice.PublisherApi.Enums.PbPrintStyle PrintStyle = 0</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget, object oneBitDownsampleThreshold, object from, object to, object copies, object collate, object printStyle, object docStructureTags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, filename, intent, includeDocumentProperties, colorDownsampleTarget, colorDownsampleThreshold, oneBitDownsampleTarget, oneBitDownsampleThreshold, from, to, copies, collate, printStyle, docStructureTags);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType Format</param>
		/// <param name="filename">string Filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="colorDownsampleTarget">optional Int32 ColorDownsampleTarget = -1</param>
		/// <param name="colorDownsampleThreshold">optional Int32 ColorDownsampleThreshold = -1</param>
		/// <param name="oneBitDownsampleTarget">optional Int32 OneBitDownsampleTarget = -1</param>
		/// <param name="oneBitDownsampleThreshold">optional Int32 OneBitDownsampleThreshold = -1</param>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="copies">optional Int32 Copies = -1</param>
		/// <param name="collate">optional bool Collate = true</param>
		/// <param name="printStyle">optional NetOffice.PublisherApi.Enums.PbPrintStyle PrintStyle = 0</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		/// <param name="bitmapMissingFonts">optional bool BitmapMissingFonts = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget, object oneBitDownsampleThreshold, object from, object to, object copies, object collate, object printStyle, object docStructureTags, object bitmapMissingFonts)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, filename, intent, includeDocumentProperties, colorDownsampleTarget, colorDownsampleThreshold, oneBitDownsampleTarget, oneBitDownsampleThreshold, from, to, copies, collate, printStyle, docStructureTags, bitmapMissingFonts);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType Format</param>
		/// <param name="filename">string Filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="colorDownsampleTarget">optional Int32 ColorDownsampleTarget = -1</param>
		/// <param name="colorDownsampleThreshold">optional Int32 ColorDownsampleThreshold = -1</param>
		/// <param name="oneBitDownsampleTarget">optional Int32 OneBitDownsampleTarget = -1</param>
		/// <param name="oneBitDownsampleThreshold">optional Int32 OneBitDownsampleThreshold = -1</param>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="copies">optional Int32 Copies = -1</param>
		/// <param name="collate">optional bool Collate = true</param>
		/// <param name="printStyle">optional NetOffice.PublisherApi.Enums.PbPrintStyle PrintStyle = 0</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		/// <param name="bitmapMissingFonts">optional bool BitmapMissingFonts = true</param>
		/// <param name="useISO19005_1">optional bool UseISO19005_1 = false</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget, object oneBitDownsampleThreshold, object from, object to, object copies, object collate, object printStyle, object docStructureTags, object bitmapMissingFonts, object useISO19005_1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, filename, intent, includeDocumentProperties, colorDownsampleTarget, colorDownsampleThreshold, oneBitDownsampleTarget, oneBitDownsampleThreshold, from, to, copies, collate, printStyle, docStructureTags, bitmapMissingFonts, useISO19005_1);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}