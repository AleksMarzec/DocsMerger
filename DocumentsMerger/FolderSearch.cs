using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentsMerger
{
    public static class FolderSearch
    {
        public static List<string> GetAllFilePaths(string rootfolderpath)
        {
            string[] temp = Directory.GetFiles(rootfolderpath, "*.odt", SearchOption.AllDirectories);

            List<string> filepaths = new List<string>(temp);
            return filepaths;
        }
    }
}