using CommandLine;
using CommandLine.Text;
using System;
using System.Collections.Generic;
using System.Reflection;

namespace SpiraAttachmentImport.Classes
{
    /// <summary>Stores the given settings for execution of the application.</summary>
    /// <remarks>https://github.com/gsscoder/commandline</remarks>
    class Settings
    {
        /// <summary>Create a new instance.</summary>
        public Settings()
        {
            ImportFilter = "*.*";
        }

        /// <summary>The server's URL.</summary>
        [Option('s', "server", Required = true, HelpText = "The URL address of the Spira install.")]
        public string SpiraServer { get; set; }

        /// <summary>The login userid to log in as.</summary>
        [Option('u', "user", Required = true, HelpText = "The Spira's user login id.")]
        public string SpiraLogin { get; set; }

        /// <summary>The password for the user to log in as.</summary>
        [Option('p', "pass", Required = true, HelpText = "The Spira's user password.")]
        public string SpiraPass { get; set; }

        /// <summary>The project ID of the project to upload to.</summary>
        [Option('t', "project", Required = true, HelpText = "The project # to import the attachments into.")]
        public long SpiraProject { get; set; }

        /// <summary>The filter to run on importing files. (Defaults *.*)</summary>
        [Option('f', "filter", DefaultValue = "*.*", HelpText = "File mask to filter imports on.")]
        public string ImportFilter { get; set; }

        /// <summary>The filter to run on importing files. (Defaults *.*)</summary>
        [Option('r', "recursive", DefaultValue = false, HelpText = "Include all subdirectories of given path.")]
        public bool PathRecursive { get; set; }

        /// <summary>The filter to run on importing files. (Defaults *.*)</summary>
        [Option('v', "verbose", DefaultValue = false, HelpText = "Debug output.")]
        public bool Verbose { get; set; }

        /// <summary>The path to pull files from.</summary>
        [ValueOption(0)]
        public string ImportPath { get; set; }

        #region Help Output
        [HelpOption]
        public string GetUsage()
        {
            var help = new HelpText
            {
                Heading = new HeadingInfo("Spira Attachment Uploader", typeof(Program).Assembly.GetName().Version.ToString()),
                Copyright = new CopyrightInfo("Inflectra Corporation", DateTime.Now.Year),
                AdditionalNewLineAfterOption = false,
                AddDashesToOption = true
            };
            help.AddPreOptionsLine("Usage: " + Environment.NewLine + "  " + typeof(Program).Assembly.GetName().Name + ".exe -s http://localhost/Spira -u administrator -p PleaseChange1 -r 2 C:\\TempUpload");
            help.AddOptions(this);
            help.MaximumDisplayWidth = 1024;
            return help;
        }
        #endregion
    }
}
