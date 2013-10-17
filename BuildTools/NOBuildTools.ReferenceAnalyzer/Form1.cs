using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using System.Text;
using WinForms = System.Windows.Forms;
using HtmlAgilityPack;

namespace NOBuildTools.ReferenceAnalyzer
{
    public partial class Form1 : WinForms.Form
    {
        public Form1()
        {
            InitializeComponent();
            XDocument document = Parser.ParseReference();
            document.Save("c:\\test123.xml");
        }
    }
}
