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
	/// DispatchInterface OMathFunction 
	/// SupportByVersion Word, 12,14,15,16
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192356.aspx
	///</summary>
	[SupportByVersionAttribute("Word", 12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class OMathFunction : COMObject
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
                    _type = typeof(OMathFunction);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public OMathFunction(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OMathFunction(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OMathFunction(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OMathFunction(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OMathFunction(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OMathFunction() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OMathFunction(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822968.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Enums.WdOMathFunctionType Type
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Type", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdOMathFunctionType)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821930.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
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
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837512.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
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
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191735.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
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
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821818.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
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
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821319.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathArgs Args
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Args", paramsArray);
				NetOffice.WordApi.OMathArgs newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.OMathArgs.LateBindingApiWrapperType) as NetOffice.WordApi.OMathArgs;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844843.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathAcc Acc
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Acc", paramsArray);
				NetOffice.WordApi.OMathAcc newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.OMathAcc.LateBindingApiWrapperType) as NetOffice.WordApi.OMathAcc;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196619.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathBar Bar
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Bar", paramsArray);
				NetOffice.WordApi.OMathBar newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.OMathBar.LateBindingApiWrapperType) as NetOffice.WordApi.OMathBar;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845538.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathBox Box
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Box", paramsArray);
				NetOffice.WordApi.OMathBox newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.OMathBox.LateBindingApiWrapperType) as NetOffice.WordApi.OMathBox;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844775.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathBorderBox BorderBox
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BorderBox", paramsArray);
				NetOffice.WordApi.OMathBorderBox newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.OMathBorderBox.LateBindingApiWrapperType) as NetOffice.WordApi.OMathBorderBox;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194269.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathDelim Delim
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Delim", paramsArray);
				NetOffice.WordApi.OMathDelim newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.OMathDelim.LateBindingApiWrapperType) as NetOffice.WordApi.OMathDelim;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840498.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathEqArray EqArray
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EqArray", paramsArray);
				NetOffice.WordApi.OMathEqArray newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.OMathEqArray.LateBindingApiWrapperType) as NetOffice.WordApi.OMathEqArray;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840155.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathFrac Frac
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Frac", paramsArray);
				NetOffice.WordApi.OMathFrac newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.OMathFrac.LateBindingApiWrapperType) as NetOffice.WordApi.OMathFrac;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821543.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathFunc Func
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Func", paramsArray);
				NetOffice.WordApi.OMathFunc newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.OMathFunc.LateBindingApiWrapperType) as NetOffice.WordApi.OMathFunc;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821338.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathGroupChar GroupChar
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "GroupChar", paramsArray);
				NetOffice.WordApi.OMathGroupChar newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.OMathGroupChar.LateBindingApiWrapperType) as NetOffice.WordApi.OMathGroupChar;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194859.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathLimLow LimLow
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LimLow", paramsArray);
				NetOffice.WordApi.OMathLimLow newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.OMathLimLow.LateBindingApiWrapperType) as NetOffice.WordApi.OMathLimLow;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838909.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathLimUpp LimUpp
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LimUpp", paramsArray);
				NetOffice.WordApi.OMathLimUpp newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.OMathLimUpp.LateBindingApiWrapperType) as NetOffice.WordApi.OMathLimUpp;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822992.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathMat Mat
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Mat", paramsArray);
				NetOffice.WordApi.OMathMat newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.OMathMat.LateBindingApiWrapperType) as NetOffice.WordApi.OMathMat;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193417.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathNary Nary
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Nary", paramsArray);
				NetOffice.WordApi.OMathNary newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.OMathNary.LateBindingApiWrapperType) as NetOffice.WordApi.OMathNary;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193856.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathPhantom Phantom
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Phantom", paramsArray);
				NetOffice.WordApi.OMathPhantom newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.OMathPhantom.LateBindingApiWrapperType) as NetOffice.WordApi.OMathPhantom;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840596.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathScrPre ScrPre
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ScrPre", paramsArray);
				NetOffice.WordApi.OMathScrPre newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.OMathScrPre.LateBindingApiWrapperType) as NetOffice.WordApi.OMathScrPre;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836951.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathRad Rad
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Rad", paramsArray);
				NetOffice.WordApi.OMathRad newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.OMathRad.LateBindingApiWrapperType) as NetOffice.WordApi.OMathRad;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840167.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathScrSub ScrSub
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ScrSub", paramsArray);
				NetOffice.WordApi.OMathScrSub newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.OMathScrSub.LateBindingApiWrapperType) as NetOffice.WordApi.OMathScrSub;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821540.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathScrSubSup ScrSubSup
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ScrSubSup", paramsArray);
				NetOffice.WordApi.OMathScrSubSup newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.OMathScrSubSup.LateBindingApiWrapperType) as NetOffice.WordApi.OMathScrSubSup;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840857.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathScrSup ScrSup
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ScrSup", paramsArray);
				NetOffice.WordApi.OMathScrSup newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.OMathScrSup.LateBindingApiWrapperType) as NetOffice.WordApi.OMathScrSup;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839640.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMath OMath
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OMath", paramsArray);
				NetOffice.WordApi.OMath newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.OMath.LateBindingApiWrapperType) as NetOffice.WordApi.OMath;
				return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196628.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathFunction Remove()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Remove", paramsArray);
			NetOffice.WordApi.OMathFunction newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.OMathFunction.LateBindingApiWrapperType) as NetOffice.WordApi.OMathFunction;
			return newObject;
		}

		#endregion
		#pragma warning restore
	}
}