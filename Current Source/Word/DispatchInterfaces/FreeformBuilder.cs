using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using LateBindingApi.Core;
namespace NetOffice.WordApi
{
	///<summary>
	/// DispatchInterface FreeformBuilder 
	/// SupportByVersion Word, 9,10,11,12,14
	///</summary>
	[SupportByVersionAttribute("Word", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class FreeformBuilder : COMObject
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
                    _type = typeof(FreeformBuilder);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FreeformBuilder(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FreeformBuilder(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FreeformBuilder(COMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FreeformBuilder() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FreeformBuilder(string progId) : base(progId)
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
				NetOffice.WordApi.Application newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Application.LateBindingApiWrapperType) as NetOffice.WordApi.Application;
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
				COMObject newObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="segmentType">NetOffice.OfficeApi.Enums.MsoSegmentType SegmentType</param>
		/// <param name="editingType">NetOffice.OfficeApi.Enums.MsoEditingType EditingType</param>
		/// <param name="x1">Single X1</param>
		/// <param name="y1">Single Y1</param>
		/// <param name="x2">optional Single X2 = 0</param>
		/// <param name="y2">optional Single Y2 = 0</param>
		/// <param name="x3">optional Single X3 = 0</param>
		/// <param name="y3">optional Single Y3 = 0</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public void AddNodes(NetOffice.OfficeApi.Enums.MsoSegmentType segmentType, NetOffice.OfficeApi.Enums.MsoEditingType editingType, Single x1, Single y1, Single x2, Single y2, Single x3, Single y3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(segmentType, editingType, x1, y1, x2, y2, x3, y3);
			Invoker.Method(this, "AddNodes", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="segmentType">NetOffice.OfficeApi.Enums.MsoSegmentType SegmentType</param>
		/// <param name="editingType">NetOffice.OfficeApi.Enums.MsoEditingType EditingType</param>
		/// <param name="x1">Single X1</param>
		/// <param name="y1">Single Y1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public void AddNodes(NetOffice.OfficeApi.Enums.MsoSegmentType segmentType, NetOffice.OfficeApi.Enums.MsoEditingType editingType, Single x1, Single y1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(segmentType, editingType, x1, y1);
			Invoker.Method(this, "AddNodes", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="segmentType">NetOffice.OfficeApi.Enums.MsoSegmentType SegmentType</param>
		/// <param name="editingType">NetOffice.OfficeApi.Enums.MsoEditingType EditingType</param>
		/// <param name="x1">Single X1</param>
		/// <param name="y1">Single Y1</param>
		/// <param name="x2">optional Single X2 = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public void AddNodes(NetOffice.OfficeApi.Enums.MsoSegmentType segmentType, NetOffice.OfficeApi.Enums.MsoEditingType editingType, Single x1, Single y1, Single x2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(segmentType, editingType, x1, y1, x2);
			Invoker.Method(this, "AddNodes", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="segmentType">NetOffice.OfficeApi.Enums.MsoSegmentType SegmentType</param>
		/// <param name="editingType">NetOffice.OfficeApi.Enums.MsoEditingType EditingType</param>
		/// <param name="x1">Single X1</param>
		/// <param name="y1">Single Y1</param>
		/// <param name="x2">optional Single X2 = 0</param>
		/// <param name="y2">optional Single Y2 = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public void AddNodes(NetOffice.OfficeApi.Enums.MsoSegmentType segmentType, NetOffice.OfficeApi.Enums.MsoEditingType editingType, Single x1, Single y1, Single x2, Single y2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(segmentType, editingType, x1, y1, x2, y2);
			Invoker.Method(this, "AddNodes", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="segmentType">NetOffice.OfficeApi.Enums.MsoSegmentType SegmentType</param>
		/// <param name="editingType">NetOffice.OfficeApi.Enums.MsoEditingType EditingType</param>
		/// <param name="x1">Single X1</param>
		/// <param name="y1">Single Y1</param>
		/// <param name="x2">optional Single X2 = 0</param>
		/// <param name="y2">optional Single Y2 = 0</param>
		/// <param name="x3">optional Single X3 = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public void AddNodes(NetOffice.OfficeApi.Enums.MsoSegmentType segmentType, NetOffice.OfficeApi.Enums.MsoEditingType editingType, Single x1, Single y1, Single x2, Single y2, Single x3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(segmentType, editingType, x1, y1, x2, y2, x3);
			Invoker.Method(this, "AddNodes", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="anchor">optional object Anchor</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public NetOffice.WordApi.Shape ConvertToShape(object anchor)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(anchor);
			object returnItem = Invoker.MethodReturn(this, "ConvertToShape", paramsArray);
			NetOffice.WordApi.Shape newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Shape.LateBindingApiWrapperType) as NetOffice.WordApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public NetOffice.WordApi.Shape ConvertToShape()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ConvertToShape", paramsArray);
			NetOffice.WordApi.Shape newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Shape.LateBindingApiWrapperType) as NetOffice.WordApi.Shape;
			return newObject;
		}

		#endregion
		#pragma warning restore
	}
}