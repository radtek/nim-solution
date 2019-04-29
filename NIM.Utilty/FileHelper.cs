using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;


namespace NIM.Utilty
{
    public class FileHelper
    {



        public static int? FindFileNameNotExistedIndex(string targetDirectoryName, string fileShortName)
        {
            if (!Directory.Exists(targetDirectoryName))
                System.IO.Directory.CreateDirectory(targetDirectoryName);

            var directoryInfo = new System.IO.DirectoryInfo(targetDirectoryName);
            var files = directoryInfo.GetFiles().ToList();


            files = files.Where(t => t.Name.ToLower().Contains(fileShortName.ToLower())).ToList();

            if (files.Count() == 0)
                return null;
            return files.Count;





        }



    }
}
