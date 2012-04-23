using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using NetOffice;
namespace NetOffice.WordApi
{
	///<summary>
	/// DispatchInterface InlineShapes 
	/// SupportByVersion Word, 9,10,11,12,14
	///</summary>
	[SupportByVersionAttribute("Word", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class InlineShapes : COMObject ,IEnumerable<NetOffice.WordApi.InlineShape>
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
                    _type = typeof(InlineShapes);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public InlineShapes(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public InlineShapes(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public InlineShapes(COMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public InlineShapes() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public InlineShapes(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public NetOffice.WordApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.WordApi.Application newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Application.LateBindingApiWrapperType) as NetOffice.WordApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public Int32 Creator
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Creator", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public Int32 Count
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Count", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="index">Int32 Index</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.WordApi.InlineShape this[Int32 index]
		{
			get
			{
				object[] paramsArray = Invoker.ValidateParamsArray(index);
				object returnItem = Invoker.MethodReturn(this, "Item", paramsArray);
				NetOffice.WordApi.InlineShape newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.InlineShape.LateBindingApiWrapperType) as NetOffice.WordApi.InlineShape;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="linkToFile">optional object LinkToFile</param>
		/// <param name="saveWithDocument">optional object SaveWithDocument</param>
		/// <param name="range">optional object Range</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public NetOffice.WordApi.InlineShape AddPicture(string fileName, object linkToFile, object saveWithDocument, object range)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, linkToFile, saveWithDocument, range);
			object returnItem = Invoker.MethodReturn(this, "AddPicture", paramsArray);
			NetOffice.WordApi.InlineShape newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.InlineShape.LateBindingApiWrapperType) as NetOffice.WordApi.InlineShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="fileName">string FileName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public NetOffice.WordApi.InlineShape AddPicture(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			object returnItem = Invoker.MethodReturn(this, "AddPicture", paramsArray);
			NetOffice.WordApi.InlineShape newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.InlineShape.LateBindingApiWrapperType) as NetOffice.WordApi.InlineShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="linkToFile">optional object LinkToFile</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public NetOffice.WordApi.InlineShape AddPicture(string fileName, object linkToFile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, linkToFile);
			object returnItem = Invoker.MethodReturn(this, "AddPicture", paramsArray);
			NetOffice.WordApi.InlineShape newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.InlineShape.LateBindingApiWrapperType) as NetOffice.WordApi.InlineShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="linkToFile">optional object LinkToFile</param>
		/// <param name="saveWithDocument">optional object SaveWithDocument</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public NetOffice.WordApi.InlineShape AddPicture(string fileName, object linkToFile, object saveWithDocument)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, linkToFile, saveWithDocument);
			object returnItem = Invoker.MethodReturn(this, "AddPicture", paramsArray);
			NetOffice.WordApi.InlineShape newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.InlineShape.LateBindingApiWrapperType) as NetOffice.WordApi.InlineShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="classType">optional object ClassType</param>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="linkToFile">optional object LinkToFile</param>
		/// <param name="displayAsIcon">optional object DisplayAsIcon</param>
		/// <param name="iconFileName">optional object IconFileName</param>
		/// <param name="iconIndex">optional object IconIndex</param>
		/// <param name="iconLabel">optional object IconLabel</param>
		/// <param name="range">optional object Range</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public NetOffice.WordApi.InlineShape AddOLEObject(object classType, object fileName, object linkToFile, object displayAsIcon, object iconFileName, object iconIndex, object iconLabel, object range)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(classType, fileName, linkToFile, displayAsIcon, iconFileName, iconIndex, iconLabel, range);
			object returnItem = Invoker.MethodReturn(this, "AddOLEObject", paramsArray);
			NetOffice.WordApi.InlineShape newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.InlineShape.LateBindingApiWrapperType) as NetOffice.WordApi.InlineShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public NetOffice.WordApi.InlineShape AddOLEObject()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "AddOLEObject", paramsArray);
			NetOffice.WordApi.InlineShape newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.InlineShape.LateBindingApiWrapperType) as NetOffice.WordApi.InlineShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="classType">optional object ClassType</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public NetOffice.WordApi.InlineShape AddOLEObject(object classType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(classType);
			object returnItem = Invoker.MethodReturn(this, "AddOLEObject", paramsArray);
			NetOffice.WordApi.InlineShape newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.InlineShape.LateBindingApiWrapperType) as NetOffice.WordApi.InlineShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="classType">optional object ClassType</param>
		/// <param name="fileName">optional object FileName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public NetOffice.WordApi.InlineShape AddOLEObject(object classType, object fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(classType, fileName);
			object returnItem = Invoker.MethodReturn(this, "AddOLEObject", paramsArray);
			NetOffice.WordApi.InlineShape newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.InlineShape.LateBindingApiWrapperType) as NetOffice.WordApi.InlineShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="classType">optional object ClassType</param>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="linkToFile">optional object LinkToFile</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public NetOffice.WordApi.InlineShape AddOLEObject(object classType, object fileName, object linkToFile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(classType, fileName, linkToFile);
			object returnItem = Invoker.MethodReturn(this, "AddOLEObject", paramsArray);
			NetOffice.WordApi.InlineShape newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.InlineShape.LateBindingApiWrapperType) as NetOffice.WordApi.InlineShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="classType">optional object ClassType</param>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="linkToFile">optional object LinkToFile</param>
		/// <param name="displayAsIcon">optional object DisplayAsIcon</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public NetOffice.WordApi.InlineShape AddOLEObject(object classType, object fileName, object linkToFile, object displayAsIcon)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(classType, fileName, linkToFile, displayAsIcon);
			object returnItem = Invoker.MethodReturn(this, "AddOLEObject", paramsArray);
			NetOffice.WordApi.InlineShape newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.InlineShape.LateBindingApiWrapperType) as NetOffice.WordApi.InlineShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="classType">optional object ClassType</param>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="linkToFile">optional object LinkToFile</param>
		/// <param name="displayAsIcon">optional object DisplayAsIcon</param>
		/// <param name="iconFileName">optional object IconFileName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public NetOffice.WordApi.InlineShape AddOLEObject(object classType, object fileName, object linkToFile, object displayAsIcon, object iconFileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(classType, fileName, linkToFile, displayAsIcon, iconFileName);
			object returnItem = Invoker.MethodReturn(this, "AddOLEObject", paramsArray);
			NetOffice.WordApi.InlineShape newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.InlineShape.LateBindingApiWrapperType) as NetOffice.WordApi.InlineShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="classType">optional object ClassType</param>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="linkToFile">optional object LinkToFile</param>
		/// <param name="displayAsIcon">optional object DisplayAsIcon</param>
		/// <param name="iconFileName">optional object IconFileName</param>
		/// <param name="iconIndex">optional object IconIndex</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public NetOffice.WordApi.InlineShape AddOLEObject(object classType, object fileName, object linkToFile, object displayAsIcon, object iconFileName, object iconIndex)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(classType, fileName, linkToFile, displayAsIcon, iconFileName, iconIndex);
			object returnItem = Invoker.MethodReturn(this, "AddOLEObject", paramsArray);
			NetOffice.WordApi.InlineShape newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.InlineShape.LateBindingApiWrapperType) as NetOffice.WordApi.InlineShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="classType">optional object ClassType</param>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="linkToFile">optional object LinkToFile</param>
		/// <param name="displayAsIcon">optional object DisplayAsIcon</param>
		/// <param name="iconFileName">optional object IconFileName</param>
		/// <param name="iconIndex">optional object IconIndex</param>
		/// <param name="iconLabel">optional object IconLabel</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public NetOffice.WordApi.InlineShape AddOLEObject(object classType, object fileName, object linkToFile, object displayAsIcon, object iconFileName, object iconIndex, object iconLabel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(classType, fileName, linkToFile, displayAsIcon, iconFileName, iconIndex, iconLabel);
			object returnItem = Invoker.MethodReturn(this, "AddOLEObject", paramsArray);
			NetOffice.WordApi.InlineShape newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.InlineShape.LateBindingApiWrapperType) as NetOffice.WordApi.InlineShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="classType">optional object ClassType</param>
		/// <param name="range">optional object Range</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public NetOffice.WordApi.InlineShape AddOLEControl(object classType, object range)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(classType, range);
			object returnItem = Invoker.MethodReturn(this, "AddOLEControl", paramsArray);
			NetOffice.WordApi.InlineShape newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.InlineShape.LateBindingApiWrapperType) as NetOffice.WordApi.InlineShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public NetOffice.WordApi.InlineShape AddOLEControl()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "AddOLEControl", paramsArray);
			NetOffice.WordApi.InlineShape newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.InlineShape.LateBindingApiWrapperType) as NetOffice.WordApi.InlineShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="classType">optional object ClassType</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public NetOffice.WordApi.InlineShape AddOLEControl(object classType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(classType);
			object returnItem = Invoker.MethodReturn(this, "AddOLEControl", paramsArray);
			NetOffice.WordApi.InlineShape newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.InlineShape.LateBindingApiWrapperType) as NetOffice.WordApi.InlineShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public NetOffice.WordApi.InlineShape New(NetOffice.WordApi.Range range)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range);
			object returnItem = Invoker.MethodReturn(this, "New", paramsArray);
			NetOffice.WordApi.InlineShape newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.InlineShape.LateBindingApiWrapperType) as NetOffice.WordApi.InlineShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="range">optional object Range</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public NetOffice.WordApi.InlineShape AddHorizontalLine(string fileName, object range)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, range);
			object returnItem = Invoker.MethodReturn(this, "AddHorizontalLine", paramsArray);
			NetOffice.WordApi.InlineShape newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.InlineShape.LateBindingApiWrapperType) as NetOffice.WordApi.InlineShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="fileName">string FileName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public NetOffice.WordApi.InlineShape AddHorizontalLine(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			object returnItem = Invoker.MethodReturn(this, "AddHorizontalLine", paramsArray);
			NetOffice.WordApi.InlineShape newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.InlineShape.LateBindingApiWrapperType) as NetOffice.WordApi.InlineShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="range">optional object Range</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public NetOffice.WordApi.InlineShape AddHorizontalLineStandard(object range)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range);
			object returnItem = Invoker.MethodReturn(this, "AddHorizontalLineStandard", paramsArray);
			NetOffice.WordApi.InlineShape newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.InlineShape.LateBindingApiWrapperType) as NetOffice.WordApi.InlineShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public NetOffice.WordApi.InlineShape AddHorizontalLineStandard()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "AddHorizontalLineStandard", paramsArray);
			NetOffice.WordApi.InlineShape newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.InlineShape.LateBindingApiWrapperType) as NetOffice.WordApi.InlineShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="range">optional object Range</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public NetOffice.WordApi.InlineShape AddPictureBullet(string fileName, object range)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, range);
			object returnItem = Invoker.MethodReturn(this, "AddPictureBullet", paramsArray);
			NetOffice.WordApi.InlineShape newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.InlineShape.LateBindingApiWrapperType) as NetOffice.WordApi.InlineShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="fileName">string FileName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public NetOffice.WordApi.InlineShape AddPictureBullet(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			object returnItem = Invoker.MethodReturn(this, "AddPictureBullet", paramsArray);
			NetOffice.WordApi.InlineShape newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.InlineShape.LateBindingApiWrapperType) as NetOffice.WordApi.InlineShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="range">optional object Range</param>
		[SupportByVersionAttribute("Word", 12,14)]
		public NetOffice.WordApi.InlineShape AddChart(NetOffice.OfficeApi.Enums.XlChartType type, object range)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, range);
			object returnItem = Invoker.MethodReturn(this, "AddChart", paramsArray);
			NetOffice.WordApi.InlineShape newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.InlineShape.LateBindingApiWrapperType) as NetOffice.WordApi.InlineShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14)]
		public NetOffice.WordApi.InlineShape AddChart()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "AddChart", paramsArray);
			NetOffice.WordApi.InlineShape newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.InlineShape.LateBindingApiWrapperType) as NetOffice.WordApi.InlineShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14)]
		public NetOffice.WordApi.InlineShape AddChart(NetOffice.OfficeApi.Enums.XlChartType type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type);
			object returnItem = Invoker.MethodReturn(this, "AddChart", paramsArray);
			NetOffice.WordApi.InlineShape newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.InlineShape.LateBindingApiWrapperType) as NetOffice.WordApi.InlineShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 14
		/// </summary>
		/// <param name="layout">NetOffice.OfficeApi.SmartArtLayout Layout</param>
		/// <param name="range">optional object Range</param>
		[SupportByVersionAttribute("Word", 14)]
		public NetOffice.WordApi.InlineShape AddSmartArt(NetOffice.OfficeApi.SmartArtLayout layout, object range)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(layout, range);
			object returnItem = Invoker.MethodReturn(this, "AddSmartArt", paramsArray);
			NetOffice.WordApi.InlineShape newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.InlineShape.LateBindingApiWrapperType) as NetOffice.WordApi.InlineShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 14
		/// </summary>
		/// <param name="layout">NetOffice.OfficeApi.SmartArtLayout Layout</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14)]
		public NetOffice.WordApi.InlineShape AddSmartArt(NetOffice.OfficeApi.SmartArtLayout layout)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(layout);
			object returnItem = Invoker.MethodReturn(this, "AddSmartArt", paramsArray);
			NetOffice.WordApi.InlineShape newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.InlineShape.LateBindingApiWrapperType) as NetOffice.WordApi.InlineShape;
			return newObject;
		}

		#endregion

       #region IEnumerable<NetOffice.WordApi.InlineShape> Member
        
        /// <summary>
		/// SupportByVersionAttribute Word, 9,10,11,12,14
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
       public IEnumerator<NetOffice.WordApi.InlineShape> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.WordApi.InlineShape item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersionAttribute Word, 9,10,11,12,14
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion
		#pragma warning restore
	}
}