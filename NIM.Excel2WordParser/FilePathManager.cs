using System;

namespace NIM.CertificationGenerator
{
    public class FilePathManager
    {

        public FilePathManager()
        {
            this.InitPaths();
        }

        private void InitPaths()
        {
            this.WordTemplateConfPath = @"c:\nim\conf";
            this.TemplateFilesPath = @"c:\nim\_temp";

            if (!System.IO.Directory.Exists(this.WordTemplateConfPath))
                throw new Exception("系统配置错误，找不到指定的WORD证书目录.");// ({this.WordTemplateConfPath}).");

            if (!System.IO.Directory.Exists(this.TemplateFilesPath))
                System.IO.Directory.CreateDirectory(this.TemplateFilesPath);

        }

        public string WordTemplateConfPath { get; private set; }

        public string TemplateFilesPath { get; private set; }

    
    }
}
