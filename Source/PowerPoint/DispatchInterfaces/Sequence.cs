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
	/// DispatchInterface Sequence 
	/// SupportByVersion PowerPoint, 10,11,12,14,15,16
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744554.aspx
	///</summary>
	[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class Sequence : Collection
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
                    _type = typeof(Sequence);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Sequence(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Sequence(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Sequence(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Sequence(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Sequence(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Sequence() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Sequence(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745955.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.PowerPointApi.Application newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.Application.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746840.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
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

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">Int32 Index</param>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.PowerPointApi.Effect this[Int32 index]
		{
			get
			{
				object[] paramsArray = Invoker.ValidateParamsArray(index);
				object returnItem = Invoker.MethodReturn(this, "Item", paramsArray);
				NetOffice.PowerPointApi.Effect newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.Effect.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Effect;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746823.aspx
		/// </summary>
		/// <param name="shape">NetOffice.PowerPointApi.Shape Shape</param>
		/// <param name="effectId">NetOffice.PowerPointApi.Enums.MsoAnimEffect effectId</param>
		/// <param name="level">optional NetOffice.PowerPointApi.Enums.MsoAnimateByLevel Level = 0</param>
		/// <param name="trigger">optional NetOffice.PowerPointApi.Enums.MsoAnimTriggerType trigger = 1</param>
		/// <param name="index">optional Int32 Index = -1</param>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Effect AddEffect(NetOffice.PowerPointApi.Shape shape, NetOffice.PowerPointApi.Enums.MsoAnimEffect effectId, object level, object trigger, object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(shape, effectId, level, trigger, index);
			object returnItem = Invoker.MethodReturn(this, "AddEffect", paramsArray);
			NetOffice.PowerPointApi.Effect newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.Effect.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Effect;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746823.aspx
		/// </summary>
		/// <param name="shape">NetOffice.PowerPointApi.Shape Shape</param>
		/// <param name="effectId">NetOffice.PowerPointApi.Enums.MsoAnimEffect effectId</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Effect AddEffect(NetOffice.PowerPointApi.Shape shape, NetOffice.PowerPointApi.Enums.MsoAnimEffect effectId)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(shape, effectId);
			object returnItem = Invoker.MethodReturn(this, "AddEffect", paramsArray);
			NetOffice.PowerPointApi.Effect newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.Effect.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Effect;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746823.aspx
		/// </summary>
		/// <param name="shape">NetOffice.PowerPointApi.Shape Shape</param>
		/// <param name="effectId">NetOffice.PowerPointApi.Enums.MsoAnimEffect effectId</param>
		/// <param name="level">optional NetOffice.PowerPointApi.Enums.MsoAnimateByLevel Level = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Effect AddEffect(NetOffice.PowerPointApi.Shape shape, NetOffice.PowerPointApi.Enums.MsoAnimEffect effectId, object level)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(shape, effectId, level);
			object returnItem = Invoker.MethodReturn(this, "AddEffect", paramsArray);
			NetOffice.PowerPointApi.Effect newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.Effect.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Effect;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746823.aspx
		/// </summary>
		/// <param name="shape">NetOffice.PowerPointApi.Shape Shape</param>
		/// <param name="effectId">NetOffice.PowerPointApi.Enums.MsoAnimEffect effectId</param>
		/// <param name="level">optional NetOffice.PowerPointApi.Enums.MsoAnimateByLevel Level = 0</param>
		/// <param name="trigger">optional NetOffice.PowerPointApi.Enums.MsoAnimTriggerType trigger = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Effect AddEffect(NetOffice.PowerPointApi.Shape shape, NetOffice.PowerPointApi.Enums.MsoAnimEffect effectId, object level, object trigger)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(shape, effectId, level, trigger);
			object returnItem = Invoker.MethodReturn(this, "AddEffect", paramsArray);
			NetOffice.PowerPointApi.Effect newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.Effect.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Effect;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745243.aspx
		/// </summary>
		/// <param name="effect">NetOffice.PowerPointApi.Effect Effect</param>
		/// <param name="index">optional Int32 Index = -1</param>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Effect Clone(NetOffice.PowerPointApi.Effect effect, object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(effect, index);
			object returnItem = Invoker.MethodReturn(this, "Clone", paramsArray);
			NetOffice.PowerPointApi.Effect newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.Effect.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Effect;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745243.aspx
		/// </summary>
		/// <param name="effect">NetOffice.PowerPointApi.Effect Effect</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Effect Clone(NetOffice.PowerPointApi.Effect effect)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(effect);
			object returnItem = Invoker.MethodReturn(this, "Clone", paramsArray);
			NetOffice.PowerPointApi.Effect newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.Effect.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Effect;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744048.aspx
		/// </summary>
		/// <param name="shape">NetOffice.PowerPointApi.Shape Shape</param>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Effect FindFirstAnimationFor(NetOffice.PowerPointApi.Shape shape)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(shape);
			object returnItem = Invoker.MethodReturn(this, "FindFirstAnimationFor", paramsArray);
			NetOffice.PowerPointApi.Effect newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.Effect.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Effect;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746508.aspx
		/// </summary>
		/// <param name="click">Int32 click</param>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Effect FindFirstAnimationForClick(Int32 click)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(click);
			object returnItem = Invoker.MethodReturn(this, "FindFirstAnimationForClick", paramsArray);
			NetOffice.PowerPointApi.Effect newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.Effect.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Effect;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746657.aspx
		/// </summary>
		/// <param name="effect">NetOffice.PowerPointApi.Effect Effect</param>
		/// <param name="level">NetOffice.PowerPointApi.Enums.MsoAnimateByLevel Level</param>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Effect ConvertToBuildLevel(NetOffice.PowerPointApi.Effect effect, NetOffice.PowerPointApi.Enums.MsoAnimateByLevel level)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(effect, level);
			object returnItem = Invoker.MethodReturn(this, "ConvertToBuildLevel", paramsArray);
			NetOffice.PowerPointApi.Effect newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.Effect.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Effect;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746103.aspx
		/// </summary>
		/// <param name="effect">NetOffice.PowerPointApi.Effect Effect</param>
		/// <param name="after">NetOffice.PowerPointApi.Enums.MsoAnimAfterEffect After</param>
		/// <param name="dimColor">optional Int32 DimColor = 0</param>
		/// <param name="dimSchemeColor">optional NetOffice.PowerPointApi.Enums.PpColorSchemeIndex DimSchemeColor = 0</param>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Effect ConvertToAfterEffect(NetOffice.PowerPointApi.Effect effect, NetOffice.PowerPointApi.Enums.MsoAnimAfterEffect after, object dimColor, object dimSchemeColor)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(effect, after, dimColor, dimSchemeColor);
			object returnItem = Invoker.MethodReturn(this, "ConvertToAfterEffect", paramsArray);
			NetOffice.PowerPointApi.Effect newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.Effect.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Effect;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746103.aspx
		/// </summary>
		/// <param name="effect">NetOffice.PowerPointApi.Effect Effect</param>
		/// <param name="after">NetOffice.PowerPointApi.Enums.MsoAnimAfterEffect After</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Effect ConvertToAfterEffect(NetOffice.PowerPointApi.Effect effect, NetOffice.PowerPointApi.Enums.MsoAnimAfterEffect after)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(effect, after);
			object returnItem = Invoker.MethodReturn(this, "ConvertToAfterEffect", paramsArray);
			NetOffice.PowerPointApi.Effect newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.Effect.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Effect;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746103.aspx
		/// </summary>
		/// <param name="effect">NetOffice.PowerPointApi.Effect Effect</param>
		/// <param name="after">NetOffice.PowerPointApi.Enums.MsoAnimAfterEffect After</param>
		/// <param name="dimColor">optional Int32 DimColor = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Effect ConvertToAfterEffect(NetOffice.PowerPointApi.Effect effect, NetOffice.PowerPointApi.Enums.MsoAnimAfterEffect after, object dimColor)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(effect, after, dimColor);
			object returnItem = Invoker.MethodReturn(this, "ConvertToAfterEffect", paramsArray);
			NetOffice.PowerPointApi.Effect newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.Effect.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Effect;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745293.aspx
		/// </summary>
		/// <param name="effect">NetOffice.PowerPointApi.Effect Effect</param>
		/// <param name="animateBackground">NetOffice.OfficeApi.Enums.MsoTriState AnimateBackground</param>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Effect ConvertToAnimateBackground(NetOffice.PowerPointApi.Effect effect, NetOffice.OfficeApi.Enums.MsoTriState animateBackground)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(effect, animateBackground);
			object returnItem = Invoker.MethodReturn(this, "ConvertToAnimateBackground", paramsArray);
			NetOffice.PowerPointApi.Effect newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.Effect.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Effect;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746429.aspx
		/// </summary>
		/// <param name="effect">NetOffice.PowerPointApi.Effect Effect</param>
		/// <param name="animateInReverse">NetOffice.OfficeApi.Enums.MsoTriState animateInReverse</param>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Effect ConvertToAnimateInReverse(NetOffice.PowerPointApi.Effect effect, NetOffice.OfficeApi.Enums.MsoTriState animateInReverse)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(effect, animateInReverse);
			object returnItem = Invoker.MethodReturn(this, "ConvertToAnimateInReverse", paramsArray);
			NetOffice.PowerPointApi.Effect newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.Effect.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Effect;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746736.aspx
		/// </summary>
		/// <param name="effect">NetOffice.PowerPointApi.Effect Effect</param>
		/// <param name="unitEffect">NetOffice.PowerPointApi.Enums.MsoAnimTextUnitEffect unitEffect</param>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Effect ConvertToTextUnitEffect(NetOffice.PowerPointApi.Effect effect, NetOffice.PowerPointApi.Enums.MsoAnimTextUnitEffect unitEffect)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(effect, unitEffect);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTextUnitEffect", paramsArray);
			NetOffice.PowerPointApi.Effect newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.Effect.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Effect;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745105.aspx
		/// </summary>
		/// <param name="pShape">NetOffice.PowerPointApi.Shape pShape</param>
		/// <param name="effectId">NetOffice.PowerPointApi.Enums.MsoAnimEffect effectId</param>
		/// <param name="trigger">NetOffice.PowerPointApi.Enums.MsoAnimTriggerType trigger</param>
		/// <param name="pTriggerShape">NetOffice.PowerPointApi.Shape pTriggerShape</param>
		/// <param name="bookmark">optional string bookmark = </param>
		/// <param name="level">optional NetOffice.PowerPointApi.Enums.MsoAnimateByLevel Level = 0</param>
		[SupportByVersionAttribute("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Effect AddTriggerEffect(NetOffice.PowerPointApi.Shape pShape, NetOffice.PowerPointApi.Enums.MsoAnimEffect effectId, NetOffice.PowerPointApi.Enums.MsoAnimTriggerType trigger, NetOffice.PowerPointApi.Shape pTriggerShape, object bookmark, object level)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pShape, effectId, trigger, pTriggerShape, bookmark, level);
			object returnItem = Invoker.MethodReturn(this, "AddTriggerEffect", paramsArray);
			NetOffice.PowerPointApi.Effect newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.Effect.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Effect;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745105.aspx
		/// </summary>
		/// <param name="pShape">NetOffice.PowerPointApi.Shape pShape</param>
		/// <param name="effectId">NetOffice.PowerPointApi.Enums.MsoAnimEffect effectId</param>
		/// <param name="trigger">NetOffice.PowerPointApi.Enums.MsoAnimTriggerType trigger</param>
		/// <param name="pTriggerShape">NetOffice.PowerPointApi.Shape pTriggerShape</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Effect AddTriggerEffect(NetOffice.PowerPointApi.Shape pShape, NetOffice.PowerPointApi.Enums.MsoAnimEffect effectId, NetOffice.PowerPointApi.Enums.MsoAnimTriggerType trigger, NetOffice.PowerPointApi.Shape pTriggerShape)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pShape, effectId, trigger, pTriggerShape);
			object returnItem = Invoker.MethodReturn(this, "AddTriggerEffect", paramsArray);
			NetOffice.PowerPointApi.Effect newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.Effect.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Effect;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745105.aspx
		/// </summary>
		/// <param name="pShape">NetOffice.PowerPointApi.Shape pShape</param>
		/// <param name="effectId">NetOffice.PowerPointApi.Enums.MsoAnimEffect effectId</param>
		/// <param name="trigger">NetOffice.PowerPointApi.Enums.MsoAnimTriggerType trigger</param>
		/// <param name="pTriggerShape">NetOffice.PowerPointApi.Shape pTriggerShape</param>
		/// <param name="bookmark">optional string bookmark = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Effect AddTriggerEffect(NetOffice.PowerPointApi.Shape pShape, NetOffice.PowerPointApi.Enums.MsoAnimEffect effectId, NetOffice.PowerPointApi.Enums.MsoAnimTriggerType trigger, NetOffice.PowerPointApi.Shape pTriggerShape, object bookmark)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pShape, effectId, trigger, pTriggerShape, bookmark);
			object returnItem = Invoker.MethodReturn(this, "AddTriggerEffect", paramsArray);
			NetOffice.PowerPointApi.Effect newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.Effect.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Effect;
			return newObject;
		}

		#endregion
		#pragma warning restore
	}
}