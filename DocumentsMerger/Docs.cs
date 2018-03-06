using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentsMerger
{
    public class Docs
    {
        List<string> InputFilesPaths { get; set; }

        public Docs()
        {
            InputFilesPaths = new List<string>();
        }

        public Docs(List<string> filepaths)
        {
            InputFilesPaths = new List<string>(filepaths);
        }

        public void AddInputFile(string path)
        {
            InputFilesPaths.Add(path);
        }

        public List<string> GetList()
        {
            return InputFilesPaths;
        }
    }
}
