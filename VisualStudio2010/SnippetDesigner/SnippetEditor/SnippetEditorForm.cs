using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.SnippetLibrary;
using Microsoft.VisualStudio.Text;
using IOleServiceProvider = Microsoft.VisualStudio.OLE.Interop.IServiceProvider;

namespace Microsoft.SnippetDesigner
{
    /// <summary>
    /// The form which represents the GUI of the snippet editor.
    /// It implements ISnippetEditor which defines what properties and functions a snippet editor
    /// must use.
    /// </summary>
    [ComVisible(true)]
    public partial class SnippetEditorForm :
        UserControl,
        ISnippetEditor
    {
        protected ILogger logger;

        //Snippy Library Access Code
        private SnippetFile snippetFile; //represents an instance of this snippet application
        private int snippetIndex; // index of the snippet in the snippetFile

        //if the user typed a single character then store it here so we know what it is
        //otherwise its null
        protected string lastCharacterEntered;

        //hash which maps the display names of the languages to the path to their user snippet directory
        internal readonly Dictionary<string, string> snippetDirectories = SnippetDirectories.Instance.UserSnippetDirectories;

        //the value of the id cell you entered before edit
        //the purpose of this is so if you modify a replcement id in the gridview we can know which ids in the codewindow
        //to update
        private string previousIDValue;


        //this is the id column value of the row you have currently selected
        //this value is use so that we know which ids to give the different higlighting
        private string currentlySelectedId;


        //store the last language selected so we know when to remove adornments
        private string previousLanguageSelected = String.Empty;

        //header snippet data
        private string snippetTitle = String.Empty;
        private string snippetDescription = String.Empty;
        private string snippetAuthor = String.Empty;
        private string snippetShortcut = String.Empty;
        private string snippetHelpUrl = String.Empty;
        private string snippetKind = String.Empty;
        private readonly List<string> snippetKeywords = new List<string>();
        private readonly List<string> snippetImports = new List<string>();
        private readonly List<string> snippetReferences = new List<string>();
        private readonly List<SnippetType> snippetTypes = new List<SnippetType>();
        public static readonly Regex ValidPotentialReplacementRegex = new Regex(StringConstants.ValidPotentialReplacementString, RegexOptions.Compiled);
        public static readonly Regex ValidExistingReplacementRegex = new Regex(StringConstants.ValidExistingReplacementString, RegexOptions.Compiled);


        protected override void OnLoad(EventArgs e)
        {
        }


        /// <summary>
        /// the current snippet we are working with in the snippet file
        /// </summary>
        public Snippet ActiveSnippet { get; set; }


        public bool FormLoadFinished { get; set; }

        private bool formDirty;

        /// <summary>
        /// Is this form in a diry state
        /// </summary>
        public bool IsFormDirty
        {
            get { return FormLoadFinished ? formDirty : false; }
            set { formDirty = value; }
        }

        /// <summary>
        /// This is the id column value of the row you have currently selected
        /// this is used for showing different markers for the active replcement
        /// </summary>
        public string CurrentlySelectedId
        {
            get { return currentlySelectedId; }
        }

        #region Properties which interact with fields in the gui

        /// <summary>
        /// File name of the snippet
        /// </summary>
        public string SnippetFileName
        {
            get { return snippetFile.FileName; }
        }

        /// <summary>
        /// Get the list of snippet titles form the codeWindowHost
        /// Set the list of items in the codeWindowHost
        /// </summary>
        public List<string> SnippetTitles
        {
            get
            {
                string[] titleArray = new string[toolStripSnippetTitles.Items.Count];
                toolStripSnippetTitles.Items.CopyTo(titleArray, 0);
                return new List<string>(titleArray);
            }
            set
            {
                toolStripSnippetTitles.Items.Clear();
                foreach (string title in value)
                {
                    toolStripSnippetTitles.Items.Add(title);
                }
                toolStripSnippetTitles.SelectedIndex = toolStripSnippetTitles.Items.IndexOf(ActiveSnippet.Title);
            }
        }


        /// <summary>
        /// The active snippet title
        /// </summary>
        public string SnippetTitle
        {
            get { return snippetTitle; }

            set { snippetTitle = value; }
        }

        public string SnippetDescription
        {
            get { return snippetDescription; }

            set
            {
                if (snippetDescription != value)
                {
                    IsFormDirty = true;
                }
                snippetDescription = value;
            }
        }

        public string SnippetAuthor
        {
            get { return snippetAuthor; }

            set
            {
                if (snippetAuthor != value)
                {
                    IsFormDirty = true;
                }
                snippetAuthor = value;
            }
        }

        public string SnippetHelpUrl
        {
            get { return snippetHelpUrl; }

            set
            {
                if (snippetHelpUrl != value)
                {
                    IsFormDirty = true;
                }
                snippetHelpUrl = value;
            }
        }

        public string SnippetShortcut
        {
            get { return snippetShortcut; }

            set
            {
                if (snippetShortcut != value)
                {
                    IsFormDirty = true;
                }
                snippetShortcut = value;
            }
        }

        public List<string> SnippetKeywords
        {
            get { return snippetKeywords; }

            set
            {
                //TODO: check if the keywords have changed
                IsFormDirty = true;


                snippetKeywords.Clear();
                snippetKeywords.AddRange(value);
            }
        }

        public string SnippetCode
        {
            get { return snippetCodeWindow.CodeText; }

            set { snippetCodeWindow.CodeText = value; }
        }

        public List<SnippetType> SnippetTypes
        {
            get { return snippetTypes; }

            set
            {
                //TODO: check if the snippettype have changed
                IsFormDirty = true;


                snippetTypes.Clear();
                snippetTypes.AddRange(value);
            }
        }

        public string SnippetKind
        {
            get { return snippetKind; }

            set { snippetKind = value; }
        }

        /// <summary>
        /// Get: Converts the snippet language from display to xml form and returns it
        /// Set: Take a string snippet language in xml form and ocnverts to display form then adds it to GUI
        /// </summary>
        public string SnippetLanguage
        {
            get
            {
                if (toolStripLanguageBox.SelectedIndex > -1)
                {
                    string langString = toolStripLanguageBox.SelectedItem.ToString();
                    if (LanguageMaps.LanguageMap.DisplayLanguageToXML.ContainsKey(langString))
                    {
                        return LanguageMaps.LanguageMap.DisplayLanguageToXML[langString];
                    }
                    else
                    {
                        return String.Empty;
                    }
                }
                else
                {
                    return String.Empty;
                }
            }

            set
            {
                string language = String.Empty;
                if (!String.IsNullOrEmpty(value) && LanguageMaps.LanguageMap.SnippetSchemaLanguageToDisplay.ContainsKey(value.ToLower()))
                {
                    language = LanguageMaps.LanguageMap.SnippetSchemaLanguageToDisplay[value.ToLower()];
                }
                else
                {
                    language = LanguageMaps.LanguageMap.ToDisplayForm(SnippetDesignerPackage.Instance.Settings.DefaultLanguage);
                }

                int index = toolStripLanguageBox.Items.IndexOf(language);
                if (index >= 0)
                {
                    toolStripLanguageBox.SelectedIndex = index;
                }
                else
                {
                    toolStripLanguageBox.SelectedIndex = 0; //select first
                }
            }
        }

        public List<string> SnippetImports
        {
            get { return snippetImports; }

            set
            {
                //This wont be called from the editor properties window 
                //so we need to figure out a different way to tell if it is in a dirty state

                snippetImports.Clear();
                snippetImports.AddRange(value);
            }
        }

        public List<string> SnippetReferences
        {
            get { return snippetReferences; }

            set
            {
                //This wont be called from the editor properties window 
                //so we need to figure out a different way to tell if it is in a dirty state

                snippetReferences.Clear();
                snippetReferences.AddRange(value);
            }
        }

        public List<Literal> SnippetReplacements
        {
            get
            {
                List<Literal> replacements = new List<Literal>();
                foreach (DataGridViewRow row in replacementGridView.Rows)
                {
                    if (row.IsNewRow) continue;
                    string currId = (string) row.Cells[StringConstants.ColumnID].EditedFormattedValue;
                    currId = currId.Trim();
                    if (String.IsNullOrEmpty(currId)) continue;

                    bool isObj = false;
                    if (((row.Cells[StringConstants.ColumnReplacementKind]).EditedFormattedValue as string).ToLower() ==
                        Resources.ReplacementObjectName.ToLower())
                    {
                        isObj = true;
                    }

                    bool isEditable = (bool) (row.Cells[StringConstants.ColumnEditable]).EditedFormattedValue;
                    replacements.Add(new Literal((string) row.Cells[StringConstants.ColumnID].EditedFormattedValue,
                                                 (string) row.Cells[StringConstants.ColumnTooltip].EditedFormattedValue,
                                                 (string) row.Cells[StringConstants.ColumnDefault].EditedFormattedValue,
                                                 (string) row.Cells[StringConstants.ColumnFunction].EditedFormattedValue,
                                                 isObj,
                                                 isEditable,
                                                 (string) row.Cells[StringConstants.ColumnType].EditedFormattedValue
                                         ));
                }
                return replacements;
            }

            set
            {
                //literals and objects
                replacementGridView.Rows.Clear();
                foreach (Literal literal in value)
                {
                    string objOrLiteral = Resources.ReplacementLiteralName;
                    if (literal.Object)
                    {
                        objOrLiteral = Resources.ReplacementObjectName;
                    }

                    object[] row = {literal.ID, literal.ToolTip, literal.DefaultValue, objOrLiteral, literal.Type, literal.Function, literal.Editable};
                    int rowIndex = replacementGridView.Rows.Add(row);
                    if (!literal.Object)
                    {
                        SetOrDisableTypeField(false, rowIndex);
                    }
                }
            }
        }

        #endregion

        /// <summary>
        /// Get the codewindow snippetExplorerForm
        /// </summary>
        public CodeWindow CodeWindow
        {
            get { return snippetCodeWindow; }
        }

        /// <summary>
        /// save snippet and update the snippets in memory
        /// </summary>
        /// <returns></returns>
        public bool SaveSnippet()
        {
            PushFieldsIntoActiveSnippet();
            if (SnippetDesignerPackage.Instance != null)
            {
                SnippetDesignerPackage.Instance.SnippetIndex.UpdateSnippetFile(snippetFile);
            }
            snippetFile.Save();
            return true;
        }

        /// <summary>
        /// Save snippet as
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public bool SaveSnippetAs(string fileName)
        {
            PushFieldsIntoActiveSnippet();
            foreach (Snippet snippetItem in snippetFile.Snippets)
            {
                if (SnippetDesignerPackage.Instance != null)
                {
                    SnippetDesignerPackage.Instance.SnippetIndex.CreateIndexItemDataFromSnippet(snippetItem, fileName);
                }
            }
            snippetFile.SaveAs(fileName);
            return true;
        }

        /// <summary>
        /// Load the snippet
        /// 
        /// throws IOExcpetion if load failure
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public bool LoadSnippet(string fileName)
        {
            try
            {
                //load the snipept into memory
                snippetFile = new SnippetFile(fileName);

                snippetIndex = 0;

                //set this snippet as the active snippet
                ActiveSnippet = snippetFile.Snippets[snippetIndex];

                //populate the gui with this snippets information
                PullFieldsFromActiveSnippet();
                //indicate that this snippet is not dirty
                IsFormDirty = false;

                FormLoadFinished = true;
            }
            catch (IOException) //abort loading snippet, fail program
            {
                //since an io error occured
                throw;
            }

            return true;
        }


        /// <summary>
        /// Takes data from in memory snippet file and populates the gui form
        /// </summary>
        public void PullFieldsFromActiveSnippet()
        {
            //snippet information
            SnippetTitle = ActiveSnippet.Title;
            SnippetAuthor = ActiveSnippet.Author;
            SnippetDescription = ActiveSnippet.Description;
            SnippetHelpUrl = ActiveSnippet.HelpUrl;
            SnippetShortcut = ActiveSnippet.Shortcut;
            SnippetKeywords = ActiveSnippet.Keywords;


            SnippetTitles = GetSnippetTitles();

            if (ActiveSnippet.SnippetTypes.Count <= 0)
            {
                //if no type specified then make it expansion by default
                snippetTypes.Add(new SnippetType(StringConstants.SnippetTypeExpansion));
            }
            else
            {
                SnippetTypes = ActiveSnippet.SnippetTypes;
            }


            //literals and objects
            SnippetReplacements = ActiveSnippet.Literals;

            //code - for some unknown reason this must be done before language is set to stop some inconsitency
            //including highlighting and color coding 
            SnippetCode = ActiveSnippet.Code;

            //kind and language values
            SnippetKind = ActiveSnippet.CodeKindAttribute;

            SnippetLanguage = ActiveSnippet.CodeLanguageAttribute;

            //imports and references
            SnippetImports = ActiveSnippet.Imports;

            SnippetReferences = ActiveSnippet.References;
        }


        /// <summary>
        /// Takes the data from the form and adds it to the in memory xml document
        /// </summary>
        public void PushFieldsIntoActiveSnippet()
        {
            //add header info
            ActiveSnippet.Title = SnippetTitle;
            ActiveSnippet.Author = SnippetAuthor;
            ActiveSnippet.Description = SnippetDescription;
            ActiveSnippet.HelpUrl = SnippetHelpUrl;
            ActiveSnippet.Shortcut = SnippetShortcut;

            //update keywords
            ActiveSnippet.Keywords = SnippetKeywords;


            //add snippet types
            ActiveSnippet.SnippetTypes = SnippetTypes;


            //add code
            ActiveSnippet.Code = SnippetCode;


            //must be after code node is declared
            //kind and language values
            ActiveSnippet.CodeKindAttribute = SnippetKind;


            ActiveSnippet.CodeLanguageAttribute = SnippetLanguage;


            //imports and references
            ActiveSnippet.Imports = SnippetImports;

            ActiveSnippet.References = SnippetReferences;

            //literals and objects
            ActiveSnippet.Literals = SnippetReplacements;
        }

        private List<String> GetSnippetTitles()
        {
            List<string> snippetTitles = new List<string>();
            foreach (Snippet s in snippetFile.Snippets)
            {
                snippetTitles.Add(s.Title);
            }
            return snippetTitles;
        }

        private void languageComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            ToolStripComboBox langCombo = sender as ToolStripComboBox;
            if (langCombo != null)
            {
                string languageText = langCombo.SelectedItem.ToString();
                if (previousLanguageSelected != languageText) //make sure this is actually a change
                {
                    IsFormDirty = true;

                    //if (languageText == Resources.DisplayNameXML)
                    //{
                    //    //The XML Editor defines its own properties window and by removing adornments it will stop it
                    //    // from showing and allow ours to show
                    //    IOleServiceProvider sp = snippetCodeWindow.VsCodeWindow as IOleServiceProvider;
                    //    if (sp != null)
                    //    {
                    //        ServiceProvider site = new ServiceProvider(sp);
                    //        IVsCodeWindowManager cMan = site.GetService(typeof(SVsCodeWindowManager)) as IVsCodeWindowManager;
                    //        if (cMan != null)
                    //        {
                    //            cMan.RemoveAdornments();
                    //        }
                    //    }
                    //}

                    //store the last language
                    previousLanguageSelected = languageText;

                    //refresh the properties window
                    SnippetEditor sEditor = (this as SnippetEditor);
                    if (sEditor != null)
                    {
                        sEditor.RefreshPropertiesWindow();
                    }
                }
            }
        }

        private void toolStripSnippetsTitles_SelectedIndexChanged(object sender, EventArgs e)
        {
            ToolStripComboBox snippetsBox = sender as ToolStripComboBox;
            if (snippetsBox != null)
            {
                string newTitle = snippetsBox.SelectedItem as string;
                if (!String.IsNullOrEmpty(newTitle) && newTitle != ActiveSnippet.Title)
                {
                    PushFieldsIntoActiveSnippet();

                    //foreach (Snippet sn in snippetFile.Snippets)
                    for (int i = 0; i < snippetFile.Snippets.Count; i++)
                    {
                        if (snippetFile.Snippets[i].Title.Equals(newTitle, StringComparison.InvariantCulture))
                        {
                            snippetIndex = i;
                            ActiveSnippet = snippetFile.Snippets[i];
                        }
                    }

                    PullFieldsFromActiveSnippet();

                    //clear and show all markers
                    //RefreshReplacementMarkers(false);

                    //not the best way to do this but since I dont know if we want to move the change current snippet to the 
                    // porperties window this will have to do for now
                    // I am assuming this object is actually an instance of snippeteditor
                    SnippetEditor theEditor = this as SnippetEditor;
                    if (theEditor != null)
                    {
                        theEditor.RefreshPropertiesWindow();
                    }
                }
            }
        }

        private void toolStripSnippetTitles_TextUpdate(object sender, EventArgs e)
        {
            ToolStripComboBox snippetsBox = sender as ToolStripComboBox;
            Debug.WriteLine("Text Update " + snippetsBox.Text);

            if (snippetsBox != null)
            {
                string newTitle = snippetsBox.Text;
                if (!String.IsNullOrEmpty(newTitle) && newTitle != snippetTitle)
                {
                    snippetsBox.Items.Remove(snippetTitle);
                    snippetTitle = newTitle;
                    PushFieldsIntoActiveSnippet();
                    IsFormDirty = true;
                }
            }
        }

        private void SetOrDisableTypeField(bool isObject, int rowIndex)
        {
            if (isObject) //if this is an object than enable the type field
            {
                replacementGridView.Rows[rowIndex].Cells[StringConstants.ColumnType].Value = String.Empty;
                replacementGridView.Rows[rowIndex].Cells[StringConstants.ColumnType].ReadOnly = false;
            }
            else //this is not a object so disable type field
            {
                replacementGridView.Rows[rowIndex].Cells[StringConstants.ColumnType].Value = Resources.TypeInvalidForLiteralSymbol;
                replacementGridView.Rows[rowIndex].Cells[StringConstants.ColumnType].ReadOnly = true;
            }
        }

        #region Replacement Grid Events

        private void replacementGridView_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            DataGridView grid = sender as DataGridView;
            if (grid != null)
            {
                if (grid.Columns[e.ColumnIndex].Name == StringConstants.ColumnID)
                {
                    previousIDValue = (string) grid.Rows[e.RowIndex].Cells[e.ColumnIndex].EditedFormattedValue;
                    return;
                }
            }
            previousIDValue = null;
        }

        private void replacementGridView_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            DataGridView grid = sender as DataGridView;
            if (grid != null)
            {
                currentlySelectedId = grid.Rows[e.RowIndex].Cells[StringConstants.ColumnID].Value as string;
                RefreshReplacementMarkers();
            }
        }

        private void replacementGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            DataGridView grid = sender as DataGridView;
            if (grid != null)
            {
                IsFormDirty = true;
                if (grid.Columns[e.ColumnIndex].Name == StringConstants.ColumnID)
                {
                    string newIdValue = (string) grid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
                    if (newIdValue == null)
                    {
                        //if null make it empty
                        newIdValue = String.Empty;
                    }

                    //make sure a change is being made
                    if (previousIDValue != null && newIdValue != previousIDValue)
                    {
                        //check if the change you made is valid if not tell the user and return to previous value
                        if (IsValidReplaceableText(newIdValue))
                        {
                            //build new replacement text
                            string newReplacement = TurnTextIntoReplacementSymbol(newIdValue);
                            string oldReplacement = TurnTextIntoReplacementSymbol(previousIDValue);

                            //replace all occurances of the oldReplacement with newReplacement
                            //set false so it allows us to override existing replacements
                            ReplaceAll(oldReplacement, newReplacement, false);
                            ReplaceAll(newIdValue, newReplacement, true);

                            IsFormDirty = true; //form is now dirty
                        }
                        else
                        {
                            AlertInvalidReplacement(newIdValue);

                            //set id cell back to the old value
                            grid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = previousIDValue;
                        }
                    }
                }
                else if (grid.Columns[e.ColumnIndex].Name == StringConstants.ColumnReplacementKind)
                {
                    if ((string) grid.Rows[e.RowIndex].Cells[e.ColumnIndex].EditedFormattedValue == Resources.ReplacementLiteralName)
                    {
                        SetOrDisableTypeField(false, e.RowIndex);
                    }
                    else
                    {
                        SetOrDisableTypeField(true, e.RowIndex);
                    }
                }
            }
        }

        private void replacementGridView_RowsRemoved(object sender, DataGridViewRowCancelEventArgs e)
        {
            UpdateMarkersAfterDeletedGridViewRow(e.Row);
        }

        private void removeReplacementToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // If there is one row then its the default row which we cant delete
            if (replacementGridView.Rows.Count > 1)
            {
                if (replacementGridView.SelectedCells.Count > 0)
                {
                    int rowIndex = replacementGridView.SelectedCells[0].RowIndex;

                    // make sure we are not deleting that last row which is the default one
                    if (rowIndex < replacementGridView.Rows.Count - 1)
                    {
                        DataGridViewRow rowToDelete = replacementGridView.Rows[rowIndex];
                        //update markers and remove the row
                        UpdateMarkersAfterDeletedGridViewRow(rowToDelete);
                        replacementGridView.Rows.Remove(rowToDelete);
                    }
                }
            }
        }

        private void snippetReplacementGrid_MouseDown(object sender, MouseEventArgs e)
        {
            DataGridView.HitTestInfo info = replacementGridView.HitTest(e.X, e.Y);
            if (info.RowIndex >= 0)
            {
                replacementGridView.Rows[info.RowIndex].Selected = true;
            }
        }

        #endregion

        private void AlertInvalidReplacement(string newIdValue)
        {
            logger.MessageBox("Invalid Replacement", String.Format(Resources.ErrorInvalidReplacementID, newIdValue), LogType.Warning);
        }

        private static string TurnTextIntoReplacementSymbol(string text)
        {
            return StringConstants.SymbolReplacement + text + StringConstants.SymbolReplacement;
        }

        private List<string> GetCurrentReplacements()
        {
            List<string> allReplacements = new List<string>();
            foreach (DataGridViewRow row in replacementGridView.Rows)
            {
                string idValue = ((string) row.Cells[StringConstants.ColumnID].EditedFormattedValue).Trim();
                if (idValue.Length > 0)
                {
                    allReplacements.Add(idValue);
                }
            }
            return allReplacements;
        }

        protected void RefreshReplacementMarkers(int lineToMark)
        {
            var allReplacements = GetCurrentReplacements();

            //search through the code window and update all replcement highlight martkers
            MarkReplacements(allReplacements, lineToMark);
        }

        protected void RefreshReplacementMarkers()
        {
            RefreshReplacementMarkers(-1);
        }

        public void MakeClickedReplacementActive()
        {
            SnapshotSpan currentWordSpan;
            //see if the person clicked inside of a replacement and return its span
            if (GetClickedOnReplacementSpan(out currentWordSpan))
            {
                string currentWord = CodeWindow.GetSpanText(currentWordSpan);

                foreach (DataGridViewRow row in replacementGridView.Rows)
                {
                    if ((string) row.Cells[StringConstants.ColumnID].Value == currentWord)
                    {
                        replacementGridView.ClearSelection();
                        row.Selected = true;
                        currentlySelectedId = row.Cells[StringConstants.ColumnID].Value as string;
                        break;
                    }
                }
            }
        }

        public void ReplacementRemove()
        {
            SnapshotSpan currentWordSpan;
            if (GetClickedOnReplacementSpan(out currentWordSpan))
            {
                string currentWord = currentWordSpan.GetText();
                ReplacementRemove(currentWord);
            }
        }

        private void ReplacementRemove(string textToChange)
        {
            DataGridViewRow rowToDelete = null;
            foreach (DataGridViewRow row in replacementGridView.Rows)
            {
                if ((string) row.Cells[StringConstants.ColumnID].Value == textToChange)
                {
                    rowToDelete = row;
                    break;
                }
            }

            if (rowToDelete != null)
            {
                UpdateMarkersAfterDeletedGridViewRow(rowToDelete);
                replacementGridView.Rows.Remove(rowToDelete);
            }
        }

        public void CreateReplacementFromSelection()
        {
            string selectedText;
            var selectionLength = CodeWindow.Selection.Length;
            if (selectionLength > 0)
            {
                //trim any replacement symbols or spaces
                selectedText = CodeWindow.Selection.GetText().Trim();
            }
            else
            {
                selectedText = CodeWindow.GetWordFromCurrentPosition();
            }


            if (!IsValidReplaceableText(selectedText))
            {
                AlertInvalidReplacement(selectedText);
            }
            else if (CreateReplacement(selectedText) && selectionLength > 0)
            {
                var newSnapshot = CodeWindow.Selection.Snapshot.TextBuffer.CurrentSnapshot;
                // CodeWindow.Selection = new SnapshotSpan(newSnapshot, new Span(CodeWindow.Selection.Start.Position, CodeWindow.Selection.End.Position + (StringConstants.SymbolReplacement.Length * 2)));
            }
        }

        private bool IsValidReplaceableText(string text)
        {
            return ValidPotentialReplacementRegex.IsMatch(text);
        }

        private bool CreateReplacement(string textToChange)
        {
            //check if replacement exists already
            bool existsAlready = false;
            foreach (DataGridViewRow row in replacementGridView.Rows)
            {
                if ((string) row.Cells[StringConstants.ColumnID].EditedFormattedValue == textToChange ||
                    textToChange.Trim() == String.Empty)
                {
                    //this replacement already exists or is nothing don't add it to the replacement list
                    existsAlready = true;
                }
            }

            //build new replacement text
            string newText = TurnTextIntoReplacementSymbol(textToChange);
            if (!existsAlready)
            {
                object[] newRow = {textToChange, textToChange, textToChange, Resources.ReplacementLiteralName, String.Empty, String.Empty, true};
                int rowIndex = replacementGridView.Rows.Add(newRow);
                SetOrDisableTypeField(false, rowIndex);
            }

            //replace all occurances of the textToFind with $textToFind$
            int numFoundAndReplaced = ReplaceAll(textToChange, newText, true);

            return numFoundAndReplaced > 0;
        }

        private void UpdateMarkersAfterDeletedGridViewRow(DataGridViewRow deletedRow)
        {
            if (deletedRow != null)
            {
                string deletedID = deletedRow.Cells[StringConstants.ColumnID].EditedFormattedValue as string;
                if (deletedID != null)
                {
                    //build new replacement text 
                    string currentText = TurnTextIntoReplacementSymbol(deletedID);
                    ReplaceAll(currentText, deletedID, false);
                }
            }
        }

        private void MarkReplacements(ICollection<string> replaceIDs, int lineToMark)
        {
            if (replaceIDs == null)
            {
                return;
            }

            int lineLength;
            int startLine = 0;
            int endLine = CodeWindow.LineCount;

            if (lineToMark > -1) //are we just replacing markers on the given line
            {
                startLine = lineToMark;
                endLine = startLine + 1;
            }

            //loop through all the lines we are searching
            for (int line = startLine; line < endLine; line++)
            {
                //get the length of this line
                lineLength = CodeWindow.LineLength(line);

                //loop over the line looking for $ and find the next matching one
                for (int index = 0; index < lineLength; index++)
                {
                    //find the character at this position
                    string character = CodeWindow.GetCharacterAtPosition(new TextPoint(line, index));
                    //check if this character is the replacement symbol
                    if (character == StringConstants.SymbolReplacement)
                    {
                        int nextIndex = index + 1;
                        while (nextIndex < lineLength && CodeWindow.GetCharacterAtPosition(new TextPoint(line, nextIndex)) != StringConstants.SymbolReplacement)
                        {
                            nextIndex++;
                        }
                        if (nextIndex < lineLength) //we found another SymbolReplacement
                        {
                            //create text span for the space between the two SnippetDesigner.StringConstants.ReplacementSymbols
                            //make sure text between the $ symbols matches replaceID
                            var lineText = CodeWindow.TextBuffer.CurrentSnapshot.GetLineFromLineNumber(line).GetText();
                            string textBetween = lineText.Substring(index + 1, nextIndex - (index + 1));

                            if (replaceIDs.Contains(textBetween))
                            {
                                index = nextIndex;
                                //skip the ending SnippetDesigner.StringConstants.SymbolReplacement, it will be incremented the one extra in the next loop iteration
                            }
                            else
                            {
                                string trimedText = textBetween.Trim();
                                //this replacement does not exist yet so create it only if the last character entered was the replacement symbol
                                if (lastCharacterEntered != null //make sure a single character was just entered
                                    && lastCharacterEntered == StringConstants.SymbolReplacement //make sure the last charcter is a $
                                    && trimedText == textBetween //make sure this replacement doesnt have whitespace in it
                                    && trimedText != String.Empty //and make sure its not empty
                                    && trimedText != StringConstants.SymbolEndWord //the word cant be end
                                    && trimedText != StringConstants.SymbolSelectedWord // and the word cant be selected they have special meaning
                                    && IsValidReplaceableText(textBetween)
                                    )
                                {
                                    //make the text into a replacement but dont add the replacement symbols since the user is doing it
                                    CreateReplacement(textBetween);
                                    RefreshReplacementMarkers(lineToMark);
                                    //clear last character 
                                    lastCharacterEntered = null;
                                }
                                else
                                {
                                    index = nextIndex - 1; //subtract one since it will be incrememented in the next loop iteration
                                }
                            }
                        }
                    }
                }
            }
        }

        private int ReplaceAll(string currentWord, string newWord, bool skipIfAlreadyReplacement)
        {
            var textview = CodeWindow.TextView;
            int numberReplaced = 0;

            var textEdit = textview.TextBuffer.CreateEdit();
            foreach (var replaceableSpan in GetReplaceableSpans(textview.TextBuffer.CurrentSnapshot, currentWord))
            {
                //replace the span with the given text
                if (ReplaceSpanWithText(textEdit, newWord, replaceableSpan, skipIfAlreadyReplacement))
                {
                    numberReplaced++;
                }
            }
            textEdit.Apply();
            return numberReplaced;
        }

        private static bool ReplaceSpanWithText(ITextEdit textEdit, string newWord, SnapshotSpan replaceSpan, bool skipIfAlreadyReplacement)
        {
            try
            {
                if ((skipIfAlreadyReplacement && !IsSpanReplacement(replaceSpan)) || !skipIfAlreadyReplacement)
                {
                    textEdit.Replace(replaceSpan.Span, newWord);
                    return true;
                }
            }
            catch (Exception)
            {
                return false;
            }

            return false;
        }

        private static bool FindEnclosingReplacementQuoteSpan(SnapshotSpan span, out SnapshotSpan quoteSpan)
        {
            var lineSnapshot = span.Snapshot.GetLineFromPosition(span.Start.Position);

            int left = span.Start.Position - 1;
            int right = span.End.Position;
            var text = span.Snapshot.GetText();
            quoteSpan = new SnapshotSpan();
            while (left >= lineSnapshot.Start.Position && text[left].ToString() != StringConstants.DoubleQuoteString)
            {
                left--;
            }
            while (right < lineSnapshot.End.Position && text[right].ToString() != StringConstants.DoubleQuoteString)
            {
                right++;
            }
            if (right >= lineSnapshot.End.Position || left < lineSnapshot.Start.Position)
            {
                return false; //we didnt find a quoted replacement string
            }

            quoteSpan = new SnapshotSpan(span.Snapshot, new Span(left, right + 1 - left));

            //is this span surrounded by the replcement markers
            if (!IsSpanReplacement(quoteSpan))
            {
                return false;
            }

            //we have a correct quotes string replcement
            return true;
        }

        private bool GetClickedOnReplacementSpan(out SnapshotSpan replacementSpan)
        {
            replacementSpan = CodeWindow.GetWordTextSpanFromCurrentPosition();
            var currentWordSpan = replacementSpan;

            string currentWord = currentWordSpan.GetText();

            if (String.IsNullOrEmpty(currentWord)) //you might have selected more than a word, so use what you selected
            {
                replacementSpan = CodeWindow.Selection;
            }

            //make sure this is infact a replacement
            if (!IsSpanReplacement(currentWordSpan))
            {
                //this span doesnt seem to be a replacement but maybe its a string and the user just
                //clicked in the middle of it so lets intelligently see if thats true
                if (!FindEnclosingReplacementQuoteSpan(currentWordSpan, out replacementSpan))
                {
                    return false;
                }
            }

            //we have found a replacement span
            return true;
        }

        private static bool IsTextReplacement(string text)
        {
            var firstChar = text[0];
            var lastChar = text[text.Length - 1];

            //check the first and last characters of the span to see if they are the replacement symbols
            if (firstChar.ToString() == StringConstants.SymbolReplacement &&
                lastChar.ToString() == StringConstants.SymbolReplacement
                )
            {
                return true;
            }

            return false;
        }

        private static bool IsSpanReplacement(SnapshotSpan replaceSpan)
        {
            var spanText = replaceSpan.GetText();
            if (string.IsNullOrEmpty(spanText))
                return false;

            if (IsTextReplacement(spanText))
                return true;

            //see if replacement symbols surround this span
            if (replaceSpan.Span.Start - 1 >= 0 && replaceSpan.Span.End < replaceSpan.Snapshot.Length)
            {
                var beforeFirstChar = replaceSpan.Snapshot[replaceSpan.Span.Start - 1];
                var afterLastChar = replaceSpan.Snapshot[replaceSpan.Span.End];

                if (beforeFirstChar.ToString() == StringConstants.SymbolReplacement &&
                    afterLastChar.ToString() == StringConstants.SymbolReplacement
                    )
                {
                    return true;
                }
            }


            return false;
        }

        private IEnumerable<SnapshotSpan> GetReplaceableSpans(ITextSnapshot textSnapshot, string textToFind)
        {
            var textView = CodeWindow.TextView;
            bool isReplacement = IsTextReplacement(textToFind);
            var text = textSnapshot.GetText();
            int start = 0;
            while ((start = text.IndexOf(textToFind, start)) != -1)
            {
                var end = start + textToFind.Length;
                if ((start - 1 < 0 || !CodeWindow.IsWordChar(text[start - 1])) && (end >= text.Length || !CodeWindow.IsWordChar(text[end]))
                    || isReplacement)
                    yield return new SnapshotSpan(textSnapshot, start, textToFind.Length);
                start += textToFind.Length;
            }
        }

        private void mainObjectsRepaiont_Paint(object sender, PaintEventArgs e)
        {
            SnippetDesignerPackage.Instance.ActiveSnippetLanguage = SnippetLanguage;
            SnippetDesignerPackage.Instance.ActiveSnippetTitle = SnippetTitle;
        }
    }
}