using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// DispatchInterface Module 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835649.aspx </remarks>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Module : COMObject
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
                    _type = typeof(Module);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Module(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Module(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Module(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Module(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Module(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Module(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Module() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Module(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197648.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Application>(this, "Application", NetOffice.AccessApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845790.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192086.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Name", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820960.aspx </remarks>
		/// <param name="line">Int32 line</param>
		/// <param name="numLines">Int32 numLines</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_Lines(Int32 line, Int32 numLines)
		{
			return Factory.ExecuteStringPropertyGet(this, "Lines", line, numLines);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Lines
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820960.aspx </remarks>
		/// <param name="line">Int32 line</param>
		/// <param name="numLines">Int32 numLines</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), Redirect("get_Lines")]
		public string Lines(Int32 line, Int32 numLines)
		{
			return get_Lines(line, numLines);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195500.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 CountOfLines
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "CountOfLines");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836419.aspx </remarks>
		/// <param name="procName">string procName</param>
		/// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 get_ProcStartLine(string procName, NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind)
		{
			return Factory.ExecuteInt32PropertyGet(this, "ProcStartLine", procName, procKind);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_ProcStartLine
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836419.aspx </remarks>
		/// <param name="procName">string procName</param>
		/// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), Redirect("get_ProcStartLine")]
		public Int32 ProcStartLine(string procName, NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind)
		{
			return get_ProcStartLine(procName, procKind);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835086.aspx </remarks>
		/// <param name="procName">string procName</param>
		/// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 get_ProcCountLines(string procName, NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind)
		{
			return Factory.ExecuteInt32PropertyGet(this, "ProcCountLines", procName, procKind);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_ProcCountLines
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835086.aspx </remarks>
		/// <param name="procName">string procName</param>
		/// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), Redirect("get_ProcCountLines")]
		public Int32 ProcCountLines(string procName, NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind)
		{
			return get_ProcCountLines(procName, procKind);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822434.aspx </remarks>
		/// <param name="procName">string procName</param>
		/// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 get_ProcBodyLine(string procName, NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind)
		{
			return Factory.ExecuteInt32PropertyGet(this, "ProcBodyLine", procName, procKind);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_ProcBodyLine
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822434.aspx </remarks>
		/// <param name="procName">string procName</param>
		/// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), Redirect("get_ProcBodyLine")]
		public Int32 ProcBodyLine(string procName, NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind)
		{
			return get_ProcBodyLine(procName, procKind);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195085.aspx </remarks>
		/// <param name="line">Int32 line</param>
		/// <param name="pprockind">NetOffice.VBIDEApi.Enums.vbext_ProcKind pprockind</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_ProcOfLine(Int32 line, out NetOffice.VBIDEApi.Enums.vbext_ProcKind pprockind)
		{		
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			pprockind = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(line, pprockind);
			object returnItem = Invoker.PropertyGet(this, "ProcOfLine", paramsArray, modifiers);
			pprockind = (NetOffice.VBIDEApi.Enums.vbext_ProcKind)paramsArray[1];
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_ProcOfLine
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195085.aspx </remarks>
		/// <param name="line">Int32 line</param>
		/// <param name="pprockind">NetOffice.VBIDEApi.Enums.vbext_ProcKind pprockind</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), Redirect("get_ProcOfLine")]
		public string ProcOfLine(Int32 line, out NetOffice.VBIDEApi.Enums.vbext_ProcKind pprockind)
		{
			return get_ProcOfLine(line, out pprockind);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837190.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 CountOfDeclarationLines
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "CountOfDeclarationLines");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835633.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Enums.AcModuleType Type
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcModuleType>(this, "Type");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845332.aspx </remarks>
		/// <param name="text">string text</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void InsertText(string text)
		{
			 Factory.ExecuteMethod(this, "InsertText", text);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845379.aspx </remarks>
		/// <param name="_string">string string</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void AddFromString(string _string)
		{
			 Factory.ExecuteMethod(this, "AddFromString", _string);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821352.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void AddFromFile(string fileName)
		{
			 Factory.ExecuteMethod(this, "AddFromFile", fileName);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194137.aspx </remarks>
		/// <param name="line">Int32 line</param>
		/// <param name="_string">string string</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void InsertLines(Int32 line, string _string)
		{
			 Factory.ExecuteMethod(this, "InsertLines", line, _string);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194301.aspx </remarks>
		/// <param name="startLine">Int32 startLine</param>
		/// <param name="count">Int32 count</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void DeleteLines(Int32 startLine, Int32 count)
		{
			 Factory.ExecuteMethod(this, "DeleteLines", startLine, count);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198276.aspx </remarks>
		/// <param name="line">Int32 line</param>
		/// <param name="_string">string string</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void ReplaceLine(Int32 line, string _string)
		{
			 Factory.ExecuteMethod(this, "ReplaceLine", line, _string);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845440.aspx </remarks>
		/// <param name="eventName">string eventName</param>
		/// <param name="objectName">string objectName</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 CreateEventProc(string eventName, string objectName)
		{
			return Factory.ExecuteInt32MethodGet(this, "CreateEventProc", eventName, objectName);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195471.aspx </remarks>
		/// <param name="target">string target</param>
		/// <param name="startLine">Int32 startLine</param>
		/// <param name="startColumn">Int32 startColumn</param>
		/// <param name="endLine">Int32 endLine</param>
		/// <param name="endColumn">Int32 endColumn</param>
		/// <param name="wholeWord">optional bool WholeWord = false</param>
		/// <param name="matchCase">optional bool MatchCase = false</param>
		/// <param name="patternSearch">optional bool PatternSearch = false</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool Find(string target, Int32 startLine, Int32 startColumn, Int32 endLine, Int32 endColumn, object wholeWord, object matchCase, object patternSearch)
		{
			return Factory.ExecuteBoolMethodGet(this, "Find", new object[]{ target, startLine, startColumn, endLine, endColumn, wholeWord, matchCase, patternSearch });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195471.aspx </remarks>
		/// <param name="target">string target</param>
		/// <param name="startLine">Int32 startLine</param>
		/// <param name="startColumn">Int32 startColumn</param>
		/// <param name="endLine">Int32 endLine</param>
		/// <param name="endColumn">Int32 endColumn</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool Find(string target, Int32 startLine, Int32 startColumn, Int32 endLine, Int32 endColumn)
		{
			return Factory.ExecuteBoolMethodGet(this, "Find", new object[]{ target, startLine, startColumn, endLine, endColumn });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195471.aspx </remarks>
		/// <param name="target">string target</param>
		/// <param name="startLine">Int32 startLine</param>
		/// <param name="startColumn">Int32 startColumn</param>
		/// <param name="endLine">Int32 endLine</param>
		/// <param name="endColumn">Int32 endColumn</param>
		/// <param name="wholeWord">optional bool WholeWord = false</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool Find(string target, Int32 startLine, Int32 startColumn, Int32 endLine, Int32 endColumn, object wholeWord)
		{
			return Factory.ExecuteBoolMethodGet(this, "Find", new object[]{ target, startLine, startColumn, endLine, endColumn, wholeWord });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195471.aspx </remarks>
		/// <param name="target">string target</param>
		/// <param name="startLine">Int32 startLine</param>
		/// <param name="startColumn">Int32 startColumn</param>
		/// <param name="endLine">Int32 endLine</param>
		/// <param name="endColumn">Int32 endColumn</param>
		/// <param name="wholeWord">optional bool WholeWord = false</param>
		/// <param name="matchCase">optional bool MatchCase = false</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool Find(string target, Int32 startLine, Int32 startColumn, Int32 endLine, Int32 endColumn, object wholeWord, object matchCase)
		{
			return Factory.ExecuteBoolMethodGet(this, "Find", new object[]{ target, startLine, startColumn, endLine, endColumn, wholeWord, matchCase });
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dispid">Int32 dispid</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 11,12,14,15,16)]
		public bool IsMemberSafe(Int32 dispid)
		{
			return Factory.ExecuteBoolMethodGet(this, "IsMemberSafe", dispid);
		}

		#endregion

		#pragma warning restore
	}
}
