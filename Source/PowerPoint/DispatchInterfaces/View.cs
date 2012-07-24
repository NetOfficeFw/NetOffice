using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.PowerPointApi
{
	///<summary>
	/// DispatchInterface View 
	/// SupportByVersion PowerPoint, 9,10,11,12,14,15
	///</summary>
	[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class View : COMObject
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
                    _type = typeof(View);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public View(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public View(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public View(COMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public View() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public View(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15)]
		public NetOffice.PowerPointApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.PowerPointApi.Application newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.Application.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				COMObject newObject = NetOffice.Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15)]
		public NetOffice.PowerPointApi.Enums.PpViewType Type
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Type", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.PowerPointApi.Enums.PpViewType)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15)]
		public Int32 Zoom
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Zoom", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Zoom", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15)]
		public object Slide
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Slide", paramsArray);
				COMObject newObject = NetOffice.Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Slide", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15)]
		public NetOffice.OfficeApi.Enums.MsoTriState DisplaySlideMiniature
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplaySlideMiniature", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.MsoTriState)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisplaySlideMiniature", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15)]
		public NetOffice.OfficeApi.Enums.MsoTriState ZoomToFit
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ZoomToFit", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.MsoTriState)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ZoomToFit", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15)]
		public NetOffice.PowerPointApi.PrintOptions PrintOptions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PrintOptions", paramsArray);
				NetOffice.PowerPointApi.PrintOptions newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.PrintOptions.LateBindingApiWrapperType) as NetOffice.PowerPointApi.PrintOptions;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14,15)]
		public NetOffice.OfficeApi.Enums.MsoTriState MediaControlsVisible
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MediaControlsVisible", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.MsoTriState)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14,15)]
		public Single MediaControlsLeft
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MediaControlsLeft", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14,15)]
		public Single MediaControlsTop
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MediaControlsTop", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14,15)]
		public Single MediaControlsWidth
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MediaControlsWidth", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14,15)]
		public Single MediaControlsHeight
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MediaControlsHeight", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15)]
		public void Paste()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Paste", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15)]
		public void GotoSlide(Int32 index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			Invoker.Method(this, "GotoSlide", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="dataType">optional NetOffice.PowerPointApi.Enums.PpPasteDataType DataType = 0</param>
		/// <param name="displayAsIcon">optional NetOffice.OfficeApi.Enums.MsoTriState DisplayAsIcon = 0</param>
		/// <param name="iconFileName">optional string IconFileName = </param>
		/// <param name="iconIndex">optional Int32 IconIndex = 0</param>
		/// <param name="iconLabel">optional string IconLabel = </param>
		/// <param name="link">optional NetOffice.OfficeApi.Enums.MsoTriState Link = 0</param>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15)]
		public void PasteSpecial(NetOffice.PowerPointApi.Enums.PpPasteDataType dataType, NetOffice.OfficeApi.Enums.MsoTriState displayAsIcon, string iconFileName, Int32 iconIndex, string iconLabel, NetOffice.OfficeApi.Enums.MsoTriState link)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dataType, displayAsIcon, iconFileName, iconIndex, iconLabel, link);
			Invoker.Method(this, "PasteSpecial", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15)]
		public void PasteSpecial()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "PasteSpecial", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="dataType">optional NetOffice.PowerPointApi.Enums.PpPasteDataType DataType = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15)]
		public void PasteSpecial(NetOffice.PowerPointApi.Enums.PpPasteDataType dataType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dataType);
			Invoker.Method(this, "PasteSpecial", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="dataType">optional NetOffice.PowerPointApi.Enums.PpPasteDataType DataType = 0</param>
		/// <param name="displayAsIcon">optional NetOffice.OfficeApi.Enums.MsoTriState DisplayAsIcon = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15)]
		public void PasteSpecial(NetOffice.PowerPointApi.Enums.PpPasteDataType dataType, NetOffice.OfficeApi.Enums.MsoTriState displayAsIcon)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dataType, displayAsIcon);
			Invoker.Method(this, "PasteSpecial", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="dataType">optional NetOffice.PowerPointApi.Enums.PpPasteDataType DataType = 0</param>
		/// <param name="displayAsIcon">optional NetOffice.OfficeApi.Enums.MsoTriState DisplayAsIcon = 0</param>
		/// <param name="iconFileName">optional string IconFileName = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15)]
		public void PasteSpecial(NetOffice.PowerPointApi.Enums.PpPasteDataType dataType, NetOffice.OfficeApi.Enums.MsoTriState displayAsIcon, string iconFileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dataType, displayAsIcon, iconFileName);
			Invoker.Method(this, "PasteSpecial", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="dataType">optional NetOffice.PowerPointApi.Enums.PpPasteDataType DataType = 0</param>
		/// <param name="displayAsIcon">optional NetOffice.OfficeApi.Enums.MsoTriState DisplayAsIcon = 0</param>
		/// <param name="iconFileName">optional string IconFileName = </param>
		/// <param name="iconIndex">optional Int32 IconIndex = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15)]
		public void PasteSpecial(NetOffice.PowerPointApi.Enums.PpPasteDataType dataType, NetOffice.OfficeApi.Enums.MsoTriState displayAsIcon, string iconFileName, Int32 iconIndex)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dataType, displayAsIcon, iconFileName, iconIndex);
			Invoker.Method(this, "PasteSpecial", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="dataType">optional NetOffice.PowerPointApi.Enums.PpPasteDataType DataType = 0</param>
		/// <param name="displayAsIcon">optional NetOffice.OfficeApi.Enums.MsoTriState DisplayAsIcon = 0</param>
		/// <param name="iconFileName">optional string IconFileName = </param>
		/// <param name="iconIndex">optional Int32 IconIndex = 0</param>
		/// <param name="iconLabel">optional string IconLabel = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15)]
		public void PasteSpecial(NetOffice.PowerPointApi.Enums.PpPasteDataType dataType, NetOffice.OfficeApi.Enums.MsoTriState displayAsIcon, string iconFileName, Int32 iconIndex, string iconLabel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dataType, displayAsIcon, iconFileName, iconIndex, iconLabel);
			Invoker.Method(this, "PasteSpecial", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="printToFile">optional string PrintToFile = </param>
		/// <param name="copies">optional Int32 Copies = 0</param>
		/// <param name="collate">optional NetOffice.OfficeApi.Enums.MsoTriState Collate = -99</param>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15)]
		public void PrintOut(Int32 from, Int32 to, string printToFile, Int32 copies, NetOffice.OfficeApi.Enums.MsoTriState collate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, printToFile, copies, collate);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15)]
		public void PrintOut()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15)]
		public void PrintOut(Int32 from)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15)]
		public void PrintOut(Int32 from, Int32 to)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="printToFile">optional string PrintToFile = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15)]
		public void PrintOut(Int32 from, Int32 to, string printToFile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, printToFile);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="printToFile">optional string PrintToFile = </param>
		/// <param name="copies">optional Int32 Copies = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15)]
		public void PrintOut(Int32 from, Int32 to, string printToFile, Int32 copies)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, printToFile, copies);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15
		/// </summary>
		/// <param name="shapeId">object ShapeId</param>
		[SupportByVersionAttribute("PowerPoint", 14,15)]
		public NetOffice.PowerPointApi.Player Player(object shapeId)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(shapeId);
			object returnItem = Invoker.MethodReturn(this, "Player", paramsArray);
			NetOffice.PowerPointApi.Player newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.Player.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Player;
			return newObject;
		}

		#endregion
		#pragma warning restore
	}
}