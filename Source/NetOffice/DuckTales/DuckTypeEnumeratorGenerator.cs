using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using NetOffice.Attributes;

namespace NetOffice.Duck
{
    internal class DuckTypeEnumeratorGenerator : IDisposable
    {
        internal DuckTypeEnumeratorGenerator(StringBuilder builder, MethodInfo[] methods, EnumeratorAttribute info)
        {
            HasMethods = null != methods && methods.Length > 0;
            if (!HasMethods)
                return;

            Builder = builder;
            Info = info;
            Builder.AppendLine(Environment.NewLine + "\t\t#region IEnumerable" + Environment.NewLine + Environment.NewLine);

            foreach (MethodInfo item in methods)
            {
                if (Info.Invoke == EnumeratorInvoke.Custom)
                    BuildCustomEnumeratorMethods(item);
                else
                    BuildEnumeratorMethods(item);
            }
        }

        private StringBuilder Builder { get; set; }

        private bool HasMethods { get; set; }
        
        private EnumeratorAttribute Info { get; set; }

        private void BuildCustomEnumeratorMethods(MethodInfo method)
        {         
            Builder.AppendLine("\t\tIEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()");
            Builder.AppendLine("\t\t{");
           
            Builder.AppendLine("\t\t\tint count = Count;");
            Builder.AppendLine("\t\t\tobject[] enumeratorObjects = new object[count];");
            Builder.AppendLine("\t\t\tfor (int i = 0; i < count; i++)");
            Builder.AppendLine("\t\t\t\tenumeratorObjects[i] = this[i + 1];" + Environment.NewLine);
            Builder.AppendLine("\t\t\tforeach (object item in enumeratorObjects)");
            Builder.AppendLine("\t\t\t\tyield return item;");
            Builder.AppendLine("\t\t}" + Environment.NewLine);

            Type genericType = method.ReturnType.GetGenericArguments()[0];

            Builder.AppendLine("\t\tpublic IEnumerator<" + genericType.FullName + "> GetEnumerator()");
            Builder.AppendLine("\t\t{");

            Builder.AppendLine("\t\t\tNetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);");
            Builder.AppendLine("\t\t\tforeach (" + genericType.FullName + " item in innerEnumerator)");
            Builder.AppendLine("\t\t\t\tyield return item;");

            Builder.AppendLine("\t\t}" + Environment.NewLine);
        }
        
        private void BuildEnumeratorMethods(MethodInfo method)
        {
            bool asMethod = Info.Invoke == EnumeratorInvoke.Method;
            Builder.AppendLine("\t\tIEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()");
            Builder.AppendLine("\t\t{");
            if (asMethod)
                Builder.AppendLine("\t\t\treturn NetOffice.Utils.GetDuckVariantEnumeratorAsMethod(this);");
            else
                Builder.AppendLine("\t\t\treturn NetOffice.Utils.GetDuckVariantEnumeratorAsProperty(this);");
            Builder.AppendLine("\t\t}" + Environment.NewLine);

            Type genericType = method.ReturnType.GetGenericArguments()[0];

            Builder.AppendLine("\t\tpublic IEnumerator<" + genericType.FullName + "> GetEnumerator()");
            Builder.AppendLine("\t\t{");

            Builder.AppendLine("\t\t\tNetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);");
            Builder.AppendLine("\t\t\tforeach (" + genericType.FullName + " item in innerEnumerator)");
            Builder.AppendLine("\t\t\t\tyield return item;");

            Builder.AppendLine("\t\t}" + Environment.NewLine);
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

        private static bool IsEnumeratorType(Type type)
        {
            return type.FullName.StartsWith("System.Collections.Generic.IEnumerator");
        }

        public void Dispose()
        {
            if (HasMethods)
                Builder.AppendLine(Environment.NewLine + "\t\t#endregion");
        }
    }
}
