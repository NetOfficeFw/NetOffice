using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using NetOffice.Attributes;

namespace NetOffice.Duck
{
    internal class DuckTypeMethodsGenerator : IDisposable
    {
        internal DuckTypeMethodsGenerator(StringBuilder builder, MethodInfo[] methods)
        {
            HasMethods = null != methods && methods.Length > 0;
            if (!HasMethods)
                return;

            Builder = builder;
            Builder.AppendLine(Environment.NewLine + "\t\t#region Methods" + Environment.NewLine + Environment.NewLine);

            foreach (MethodInfo item in methods)
            {
                string redirectMethodName = IsRedirectMethod(item);
                if (null != redirectMethodName)
                    RedirectMethod(item, redirectMethodName);
                else
                    BuildMethod(item);
            }
        }

        private StringBuilder Builder { get; set; }

        private bool HasMethods { get; set; }

        private void RedirectMethod(MethodInfo method, string redirectMethodName)
        {
            bool hasReturn = method.ReturnType.FullName != "System.Void";
            ParameterInfo[] arguments = method.GetParameters();

            Builder.Append("\t\tpublic " + (hasReturn ? method.ReturnType.FullName : "void") + " " + method.Name + "(");
            for (int i = 0; i < arguments.Length; i++)
            {
                ParameterInfo argument = arguments[i];
                Builder.Append(argument.ParameterType.FullName + " " + argument.Name);
                if (i < arguments.Length - 1)
                    Builder.Append(", ");
            }
            Builder.AppendLine(")");
            Builder.AppendLine("\t\t{");

            if (hasReturn)
                Builder.Append("\t\t\treturn " + redirectMethodName + "(");
            else
                Builder.Append("\t\t\t" + redirectMethodName + "(");

            for (int i = 0; i < arguments.Length; i++)
            {
                ParameterInfo argument = arguments[i];
                Builder.Append(argument.Name);
                if (i < arguments.Length - 1)
                    Builder.Append(", ");
            }
            Builder.AppendLine(");");

            Builder.AppendLine("\t\t}" + Environment.NewLine);
        }
        
        private void BuildMethod(MethodInfo method)
        {
            bool hasReturn = method.ReturnType.FullName != "System.Void";
            ParameterInfo[] arguments = method.GetParameters();
            Builder.Append("\t\tpublic " + (hasReturn ? method.ReturnType.FullName : "void") +  " " + method.Name +  "(");
            for (int i = 0; i < arguments.Length; i++)
            {
                ParameterInfo argument = arguments[i];
                Builder.Append(argument.ParameterType.FullName + " " + argument.Name);
                if (i < arguments.Length - 1)
                    Builder.Append(", ");
            }
            Builder.AppendLine(")");
            Builder.AppendLine("\t\t{");

            if (arguments.Length > 0)
            {
                Builder.Append("\t\t\tobject[] paramsArray = Invoker.ValidateParamsArray(");
                for (int i = 0; i < arguments.Length; i++)
                {
                    ParameterInfo argument = arguments[i];
                    Builder.Append(argument.Name);
                    if (i < arguments.Length - 1)
                        Builder.Append(", ");
                }
                Builder.AppendLine(");");
            }
            else
                Builder.AppendLine("\t\t\tobject[] paramsArray = null;");

            if (hasReturn)
            {
                if (InvokeAsProperty(method))
                {
                    string validatedMethodName = ValidateMethodName(method.Name);
                    Builder.AppendLine("\t\t\tobject returnItem = Invoker.PropertyGet(this, \"" + validatedMethodName + "\", paramsArray);");
                }
                else
                    Builder.AppendLine("\t\t\tobject returnItem = Invoker.MethodReturn(this, \"" + method.Name + "\", paramsArray);");

                if (IsEnumType(method.ReturnType))
                {
                    Builder.AppendLine("\t\t\tint intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);");
                    Builder.AppendLine("\t\t\t return(" + method.ReturnType.FullName + ")intReturnItem;");
                }
                if (IsSystemValueType(method.ReturnType))
                {
                    Builder.AppendLine("\t\t\treturn NetRuntimeSystem.Convert.To" + method.ReturnType.Name + "(returnItem);");
                }
                else if (IsNetOfficeReferenceType(method.ReturnType))
                {
                    Builder.AppendLine("\t\t\treturn Factory.CreateDuckObjectFromComProxy(this, returnItem, typeof(" + method.ReturnType.FullName + ")) as " + method.ReturnType.FullName + ";");
                }
                else if (IsObjectType(method.ReturnType))
                {
                    Builder.AppendLine("\t\t\tif((null != returnItem) && (returnItem is MarshalByRefObject))");
                    Builder.AppendLine("\t\t\t{");
                    Builder.AppendLine("\t\t\t\tICOMObject newObject = Factory.CreateDuckObjectFromComProxy(this, returnItem);");
                    Builder.AppendLine("\t\t\t\treturn newObject;");
                    Builder.AppendLine("\t\t\t}");
                    Builder.AppendLine("\t\t\telse");
                    Builder.AppendLine("\t\t\t{");
                    Builder.AppendLine("\t\t\t\treturn returnItem;");
                    Builder.AppendLine("\t\t\t}");
                }
                else
                    throw new InvalidProgramException();

            }
            else
            {
                if (InvokeAsProperty(method))
                {
                    string validatedMethodName = ValidateMethodName(method.Name);
                    Builder.AppendLine("\t\t\tInvoker.PropertySet(this, \"" + validatedMethodName + "\", paramsArray);");
                }
                else
                    Builder.AppendLine("\t\t\tInvoker.Method(this, \"" + method.Name + "\", paramsArray);");
            }

            Builder.AppendLine("\t\t}" + Environment.NewLine);
        }

        private string ValidateMethodName(string name)
        {
            if (name.StartsWith("get_") || name.StartsWith("set_") || name.StartsWith("add_"))
                return name.Substring(4);
            else if (name.StartsWith("remove_"))
                return name.Substring(7);
            else
                return name;
        }

        private static string IsRedirectMethod(MethodInfo method)
        {
            object[] attributes = method.GetCustomAttributes(typeof(RedirectAttribute), false);
            if (attributes.Length == 0)
                return null;
            RedirectAttribute attribute = (RedirectAttribute)attributes[0];
            return attribute.Value;
        }

        private static bool InvokeAsProperty(MethodInfo method)
        {
            object[] attributes = method.GetCustomAttributes(typeof(InvokeAsAttribute), false);
            if (attributes.Length == 0)
                return false;
            InvokeAsAttribute attribute = (InvokeAsAttribute)attributes[0];
            return attribute.Invoke == Invoke.Property;
        }

        private static bool IsObjectType(Type type)
        {         
            return type.FullName == "System.Object";
        }

        private static bool IsNetOfficeReferenceType(Type type)
        {
            return type.IsValueType == false && type.FullName.StartsWith("NetOffice");
        }

        private static bool IsSystemValueType(Type type)
        {
            return type.IsValueType || type.FullName == "System.String";
        }

        private static bool IsEnumType(Type type)
        {
            return type.IsEnum;
        }
        
        public void Dispose()
        {
            if (HasMethods)
                Builder.AppendLine(Environment.NewLine + "\t\t#endregion");
        }
    }
}
