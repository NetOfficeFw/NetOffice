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
	/// DispatchInterface Table 
	/// SupportByVersion PowerPoint, 9,10,11,12,14,15
	///</summary>
	[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class Table : COMObject
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
                    _type = typeof(Table);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Table(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Table(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Table(COMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Table() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Table(string progId) : base(progId)
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
		public NetOffice.PowerPointApi.Columns Columns
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Columns", paramsArray);
				NetOffice.PowerPointApi.Columns newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.Columns.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Columns;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15)]
		public NetOffice.PowerPointApi.Rows Rows
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Rows", paramsArray);
				NetOffice.PowerPointApi.Rows newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.Rows.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Rows;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15)]
		public NetOffice.PowerPointApi.Enums.PpDirection TableDirection
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TableDirection", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.PowerPointApi.Enums.PpDirection)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "TableDirection", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14,15)]
		public bool FirstRow
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FirstRow", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FirstRow", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14,15)]
		public bool LastRow
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LastRow", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "LastRow", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14,15)]
		public bool FirstCol
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FirstCol", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FirstCol", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14,15)]
		public bool LastCol
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LastCol", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "LastCol", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14,15)]
		public bool HorizBanding
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HorizBanding", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "HorizBanding", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14,15)]
		public bool VertBanding
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "VertBanding", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "VertBanding", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14,15)]
		public NetOffice.PowerPointApi.TableStyle Style
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Style", paramsArray);
				NetOffice.PowerPointApi.TableStyle newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.TableStyle.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TableStyle;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14,15)]
		public NetOffice.PowerPointApi.TableBackground Background
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Background", paramsArray);
				NetOffice.PowerPointApi.TableBackground newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.TableBackground.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TableBackground;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14,15)]
		public string AlternativeText
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AlternativeText", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "AlternativeText", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14,15)]
		public string Title
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Title", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Title", paramsArray);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="row">Int32 Row</param>
		/// <param name="column">Int32 Column</param>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15)]
		public NetOffice.PowerPointApi.Cell Cell(Int32 row, Int32 column)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(row, column);
			object returnItem = Invoker.MethodReturn(this, "Cell", paramsArray);
			NetOffice.PowerPointApi.Cell newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.Cell.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Cell;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15)]
		public void MergeBorders()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "MergeBorders", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15
		/// </summary>
		/// <param name="scale">Single scale</param>
		[SupportByVersionAttribute("PowerPoint", 12,14,15)]
		public void ScaleProportionally(Single scale)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(scale);
			Invoker.Method(this, "ScaleProportionally", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15
		/// </summary>
		/// <param name="styleID">optional string StyleID = </param>
		/// <param name="saveFormatting">optional bool SaveFormatting = false</param>
		[SupportByVersionAttribute("PowerPoint", 12,14,15)]
		public void ApplyStyle(object styleID, object saveFormatting)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(styleID, saveFormatting);
			Invoker.Method(this, "ApplyStyle", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 12,14,15)]
		public void ApplyStyle()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ApplyStyle", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15
		/// </summary>
		/// <param name="styleID">optional string StyleID = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 12,14,15)]
		public void ApplyStyle(object styleID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(styleID);
			Invoker.Method(this, "ApplyStyle", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}