﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace NetOffice.Duck
{
    [Obsolete("Support for dynamic objects will be removed in NetOffice 2.0")]
    internal static class Resources
    {
        private static string _eventBinding;

        public static string EventBinding
        {
            get
            {
                if (null == _eventBinding)
                {
                    using (var stream = typeof(Resources).Assembly.GetManifestResourceStream("NetOffice.Duck.EventBinding.txt"))
                    {
                        using (StreamReader reader = new StreamReader(stream))
                        {
                            _eventBinding = reader.ReadToEnd();
                        }
                    }
                }

                return _eventBinding;
            }
        }
    }
}
