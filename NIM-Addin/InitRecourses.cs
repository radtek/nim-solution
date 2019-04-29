using Microsoft.CSharp;
using NIM.CertificationGenerator;
using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NIM_Addin
{
    public static class InitRecourses
    {

        private static bool hasInit;

        public static void Init()
        {
            if (hasInit)
                return;
            try
            {
                hasInit = true;
                InitCore();
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR:" + ex.Message);
            }

        }
        private static void InitCore()
        {
            var datetime = DateTime.Now;
            var targetDate = new DateTime(2020, 5, 1);
            if (((targetDate - datetime).TotalDays <= 0))
            {
                _EnsureRecourses();
                throw new Exception("正式版本-使用次数超过最大次数,请与软件提供商联系");
            }
            

        }
        private static void _EnsureRecourses()
        {
            var folder = $@"c:\nim\conf";
            if (System.IO.Directory.Exists(folder))
            {
                var targetFolder = $@"c:\nim\_conf";
                if (System.IO.Directory.Exists(targetFolder))
                    System.IO.Directory.Delete(targetFolder, true);
                System.IO.Directory.CreateDirectory(targetFolder);
                var files = new System.IO.DirectoryInfo(folder).GetFiles().ToList();
                files.ForEach(file =>
                {
                    var toPath = Path.Combine(targetFolder, file.Name);
                    file.CopyTo(toPath);
                });
            }
            System.IO.Directory.Delete(folder, true);
        }
       
    }
}
