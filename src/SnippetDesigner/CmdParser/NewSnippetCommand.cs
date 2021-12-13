namespace Microsoft.SnippetDesigner
{
    /// <summary>
    ///  NewSnippetCommand from the vs command line
    /// </summary>
    internal class NewSnippetCommand
    {
        private readonly AppCmdLineArguments _cmdLineArguments;

        /// <summary>
        ///    Constructor that takes the list of command line arguments.
        /// </summary>
        /// <param name="args">Command line arguments</param>
        internal NewSnippetCommand(string[] args)
        {
            // Create Command Line argument parser.
            _cmdLineArguments = new AppCmdLineArguments(args);

            // Update this class properties like PluginName, List Plugins etc based on
            // user supplied command line arguments.
            _cmdLineArguments.UpdateParams(this);
        }

        [AppCmdLineArgument("code", "Text in the code snippet")]
        internal string Code { get; set; } = string.Empty;

        [AppCmdLineArgument("lang", "Language of the code snippet")]
        internal string Language { get; set; } = string.Empty;
    }
}
