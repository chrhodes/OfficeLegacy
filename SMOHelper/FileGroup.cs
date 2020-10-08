using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SMO = Microsoft.SqlServer.Management.Smo;

namespace SMOHelper
{
    public class FileGroup
    {
        private SMO.FileGroup _FileGroup;

        public string IsDefault;
        public string Name;

        /// <summary>
        /// Initializes a new instance of the FileGroup class.
        /// </summary>
        /// <param name=""></param>
        public FileGroup(SMO.FileGroup fileGroup)
        {
            _FileGroup = fileGroup;

            Name = fileGroup.Name;
            IsDefault = fileGroup.IsDefault.ToString();

        }

        private Dictionary<string, DataFile> _DataFiles;
        public Dictionary<string, DataFile> DataFiles
        {
            get
            {
                if (null == _DataFiles)
                {
                    _DataFiles = new Dictionary<string, DataFile>();

                    foreach (SMO.DataFile dataFile in _FileGroup.Files)
                    {
                        SMOHelper.DataFile df = new SMOHelper.DataFile(dataFile);
                        _DataFiles.Add(dataFile.Name, df);
                    }
                }

                return _DataFiles;
            }

            set
            {
                _DataFiles = value;
            }
        }
    }
}
