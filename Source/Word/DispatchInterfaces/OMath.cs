using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.WordApi
{
	///<summary>
	/// DispatchInterface OMath 
	/// SupportByVersion Word, 12,14,15
	///</summary>
	[SupportByVersionAttribute("Word", 12,14,15)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class OMath : COMObject
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
                    _type = typeof(OMath);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OMath(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OMath(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OMath(COMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OMath() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OMath(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15)]
		public NetOffice.WordApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.WordApi.Application newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Application.LateBindingApiWrapperType) as NetOffice.WordApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15)]
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
		/// SupportByVersion Word 12, 14, 15
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				COMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15)]
		public NetOffice.WordApi.Range Range
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Range", paramsArray);
				NetOffice.WordApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15)]
		public NetOffice.WordApi.OMathFunctions Functions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Functions", paramsArray);
				NetOffice.WordApi.OMathFunctions newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.OMathFunctions.LateBindingApiWrapperType) as NetOffice.WordApi.OMathFunctions;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15)]
		public NetOffice.WordApi.Enums.WdOMathType Type
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Type", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdOMathType)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Type", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15)]
		public NetOffice.WordApi.OMath ParentOMath
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ParentOMath", paramsArray);
				NetOffice.WordApi.OMath newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.OMath.LateBindingApiWrapperType) as NetOffice.WordApi.OMath;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15)]
		public NetOffice.WordApi.OMathFunction ParentFunction
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ParentFunction", paramsArray);
				NetOffice.WordApi.OMathFunction newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.OMathFunction.LateBindingApiWrapperType) as NetOffice.WordApi.OMathFunction;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15)]
		public NetOffice.WordApi.OMathMatRow ParentRow
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ParentRow", paramsArray);
				NetOffice.WordApi.OMathMatRow newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.OMathMatRow.LateBindingApiWrapperType) as NetOffice.WordApi.OMathMatRow;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15)]
		public NetOffice.WordApi.OMathMatCol ParentCol
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ParentCol", paramsArray);
				NetOffice.WordApi.OMathMatCol newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.OMathMatCol.LateBindingApiWrapperType) as NetOffice.WordApi.OMathMatCol;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15)]
		public NetOffice.WordApi.OMath ParentArg
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ParentArg", paramsArray);
				NetOffice.WordApi.OMath newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.OMath.LateBindingApiWrapperType) as NetOffice.WordApi.OMath;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15)]
		public Int32 ArgIndex
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ArgIndex", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15)]
		public Int32 NestingLevel
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "NestingLevel", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15)]
		public Int32 ArgSize
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ArgSize", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ArgSize", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15)]
		public NetOffice.WordApi.OMathBreaks Breaks
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Breaks", paramsArray);
				NetOffice.WordApi.OMathBreaks newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.OMathBreaks.LateBindingApiWrapperType) as NetOffice.WordApi.OMathBreaks;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15)]
		public NetOffice.WordApi.Enums.WdOMathJc Justification
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Justification", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdOMathJc)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Justification", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15)]
		public Int32 AlignPoint
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AlignPoint", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "AlignPoint", paramsArray);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15)]
		public void Linearize()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Linearize", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15)]
		public void BuildUp()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "BuildUp", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15)]
		public void Remove()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Remove", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15)]
		public void ConvertToMathText()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ConvertToMathText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15)]
		public void ConvertToNormalText()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ConvertToNormalText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15)]
		public void ConvertToLiteralText()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ConvertToLiteralText", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}