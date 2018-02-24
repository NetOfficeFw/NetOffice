using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using NetOffice;
using NetOffice.Tools;
using Vbe = NetOffice.VBIDEApi;
using NetOffice.VBIDEApi.Tools;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;

namespace VbeAddin
{
    public class ProjectCollector
    {
        public ProjectCollector(Vbe.VBE environment)
        {
            Result = new List<string>();
            var projects = environment.VBProjects;
            foreach (var item in projects)
                Result.Add(item.Name);
            projects.Dispose();
        }

        public ProjectCollector(VbaProjectRepository repository)
        {
            Result = new List<string>();
            var projects = repository.ProjectsAvailable();
            foreach (var item in projects)
                Result.Add(item);
        }

        public List<string> Result { get; private set; }
    }
}