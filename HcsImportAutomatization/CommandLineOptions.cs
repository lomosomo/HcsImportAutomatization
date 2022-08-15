using System;
using System.Collections.Generic;
using System.Text;
using CommandLine;

namespace HcsImportAutomatization
{
    public class CommandLineOptions
    {
        [Option ('d', "DryRun", HelpText = "If set the import is not actually performed")]
        public bool DryRun { get; set; }

        [Option('f', "FileToIimport", HelpText = "Full path to the file that should be imported", Required = false)]
        public string FileToImport { get; set; }
    }
}
