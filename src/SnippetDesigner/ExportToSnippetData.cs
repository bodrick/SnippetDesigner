using System;
using System.Collections.Generic;

namespace Microsoft.SnippetDesigner
{
    /// <summary>
    /// Represents the data that is exported from a code file to the snippet codeWindowHost
    /// </summary>
    public class ExportToSnippetData
    {
        //member variables
        private readonly Dictionary<string, string> _exportNameToSchemaName = new(StringComparer.OrdinalIgnoreCase);

        internal ExportToSnippetData(string code, string language)
        {
            _exportNameToSchemaName[StringConstants.ExportNameCPP] = StringConstants.SchemaNameCPP;
            _exportNameToSchemaName[StringConstants.ExportNameCSharp] = StringConstants.SchemaNameCSharp;
            _exportNameToSchemaName[StringConstants.ExportNameVisualBasic] = StringConstants.SchemaNameVisualBasic;
            _exportNameToSchemaName[StringConstants.ExportNameXML] = StringConstants.SchemaNameXML;
            _exportNameToSchemaName[StringConstants.ExportNameJavaScript] = StringConstants.SchemaNameJavaScript;
            _exportNameToSchemaName[StringConstants.ExportNameJavaScript2] = StringConstants.SchemaNameJavaScript;
            _exportNameToSchemaName[StringConstants.ExportNameHTML] = StringConstants.SchemaNameHTML;
            _exportNameToSchemaName[StringConstants.ExportNameSQL] = StringConstants.SchemaNameSQL;
            _exportNameToSchemaName[StringConstants.ExportNameSQL2] = StringConstants.SchemaNameSQL;

            Code = code;
            Language = _exportNameToSchemaName.ContainsKey(language) ? _exportNameToSchemaName[language] : string.Empty;
        }

        /// <summary>
        /// Gets the code.
        /// </summary>
        /// <value>The code.</value>
        internal string Code { get; private set; }

        /// <summary>
        /// Gets the language.
        /// </summary>
        /// <value>The language.</value>
        internal string Language { get; private set; }
    }
}
