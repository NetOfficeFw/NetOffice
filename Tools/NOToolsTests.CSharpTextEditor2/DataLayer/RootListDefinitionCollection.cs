using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOToolsTests.CSharpTextEditor2.DataLayer
{
    public class RootListDefinitionCollection : IEnumerable<RootListDefinition>
    {
        public RootListDefinitionCollection(string[] tableNames)
        {
            List = new List<RootListDefinition>();
            foreach (var item in tableNames)
                List.Add(new RootListDefinition(item));
        }

        private List<RootListDefinition> List { get; set; }

        public RootListDefinition this[string tableName]
        {
            get
            {
                foreach (var item in List)
                    if (item.Name.Equals(tableName, StringComparison.InvariantCultureIgnoreCase))
                        return item;
                throw new ArgumentOutOfRangeException(tableName);
            }
        }

        public IEnumerator<RootListDefinition> GetEnumerator()
        {
            return List.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return List.GetEnumerator();
        }
    }
}
