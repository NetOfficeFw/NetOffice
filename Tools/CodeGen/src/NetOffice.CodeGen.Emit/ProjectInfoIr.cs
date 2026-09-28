using System;
using System.Collections.Generic;
using System.Linq;

namespace NetOffice.CodeGen.Emit
{
    /// <summary>Exact contract facets for emitter-owned IFactoryInfo members.</summary>
    public sealed class WrapperProjectInfoMemberContracts
    {
        public WrapperRuntimeMemberContract Constructor { get; set; }
        public WrapperRuntimeMemberContract AssemblyName { get; set; }
        public WrapperRuntimeMemberContract AssemblyNamespace { get; set; }
        public WrapperRuntimeMemberContract ComponentGuid { get; set; }
        public WrapperRuntimeMemberContract AssemblyAttribute { get; set; }
        public WrapperRuntimeMemberContract Assembly { get; set; }
        public WrapperRuntimeMemberContract Dependencies { get; set; }
        public WrapperRuntimeMemberContract ContainsType { get; set; }
        public WrapperRuntimeMemberContract ContainsClassName { get; set; }

        internal WrapperProjectInfoMemberContracts Canonicalize()
        {
            return new WrapperProjectInfoMemberContracts
            {
                Constructor = Canonicalize(Constructor),
                AssemblyName = Canonicalize(AssemblyName),
                AssemblyNamespace = Canonicalize(AssemblyNamespace),
                ComponentGuid = Canonicalize(ComponentGuid),
                AssemblyAttribute = Canonicalize(AssemblyAttribute),
                Assembly = Canonicalize(Assembly),
                Dependencies = Canonicalize(Dependencies),
                ContainsType = Canonicalize(ContainsType),
                ContainsClassName = Canonicalize(ContainsClassName)
            };
        }

        internal IEnumerable<WrapperRuntimeMemberContract> All()
        {
            return new[] { Constructor, AssemblyName, AssemblyNamespace, ComponentGuid, AssemblyAttribute, Assembly, Dependencies, ContainsType, ContainsClassName }.Where(value => value != null);
        }

        private static WrapperRuntimeMemberContract Canonicalize(WrapperRuntimeMemberContract value) => value == null ? null : value.Canonicalize();
    }

    /// <summary>Typed clean-room plan for the generated NetOffice IFactoryInfo utility.</summary>
    public sealed class WrapperProjectInfo
    {
        public string InterfaceType { get; set; } = "IFactoryInfo";
        public string AssemblyAttributeType { get; set; } = "NetOfficeAssemblyAttribute";
        public string AssemblyNamespace { get; set; }
        public IList<string> ComponentGuids { get; set; } = new List<string>();
        public IList<string> Dependencies { get; set; } = new List<string>();
        public WrapperProjectInfoMemberContracts Contracts { get; set; }

        internal WrapperProjectInfo Canonicalize()
        {
            if (string.IsNullOrWhiteSpace(AssemblyNamespace)) throw new ArgumentException("WrapperProjectInfo.AssemblyNamespace is required.");
            var result = new WrapperProjectInfo
            {
                InterfaceType = Normalize(InterfaceType, "IFactoryInfo"),
                AssemblyAttributeType = Normalize(AssemblyAttributeType, "NetOfficeAssemblyAttribute"),
                AssemblyNamespace = AssemblyNamespace.Trim(),
                ComponentGuids = NormalizeValues(ComponentGuids),
                Dependencies = NormalizeValues(Dependencies),
                Contracts = Contracts == null ? null : Contracts.Canonicalize()
            };
            foreach (var value in result.ComponentGuids)
            {
                Guid parsed;
                if (!Guid.TryParse(value, out parsed)) throw new ArgumentException("Invalid ProjectInfo component GUID: " + value);
            }
            return result;
        }

        private static IList<string> NormalizeValues(IEnumerable<string> values)
        {
            return (values ?? Enumerable.Empty<string>()).Select(value => (value ?? string.Empty).Trim()).Where(value => value.Length != 0).ToList();
        }

        private static string Normalize(string value, string fallback) => string.IsNullOrWhiteSpace(value) ? fallback : value.Trim();
    }
}
