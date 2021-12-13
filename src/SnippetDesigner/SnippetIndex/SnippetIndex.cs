using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Serialization;
using Microsoft.SnippetDesigner.SnippetExplorer;
using SnippetLibrary;

namespace Microsoft.SnippetDesigner
{
    /// <summary>
    /// Represents the index file of snippets
    /// </summary>
    public class SnippetIndex : INotifyPropertyChanged
    {
        private readonly List<Tuple<Func<SnippetIndexItem, string>, int>> _fieldRankings =
            new()
            {
                Tuple.Create<Func<SnippetIndexItem, string>, int>(snippet => snippet.Title, 10),
                Tuple.Create<Func<SnippetIndexItem, string>, int>(snippet => snippet.Code, 5),
                Tuple.Create<Func<SnippetIndexItem, string>, int>(snippet => snippet.Description, 3),
                Tuple.Create<Func<SnippetIndexItem, string>, int>(snippet => snippet.Keywords, 2),
                Tuple.Create<Func<SnippetIndexItem, string>, int>(snippet => Path.GetFileNameWithoutExtension(snippet.File), 2)
            };

        // Maps SnippetFilePath|SnippetTitle to SnippetIndexItem
        private readonly Dictionary<string, SnippetIndexItem> _indexedSnippets;

        private readonly ILogger _logger;
        private readonly string _snippetIndexFilePath;
        private bool _isIndexLoading;
        private bool _isIndexUpdating;

        public SnippetIndex()
        {
            _snippetIndexFilePath = SnippetDesignerPackage.Instance.Settings.SnippetIndexLocation;
            _indexedSnippets = new Dictionary<string, SnippetIndexItem>();
            _logger = SnippetDesignerPackage.Instance.Logger;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// Gets or sets a value indicating whether this instance is index loading.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if this instance is index loading; otherwise, <c>false</c>.
        /// </value>
        public bool IsIndexLoading
        {
            get => _isIndexLoading;
            set
            {
                if (_isIndexLoading != value)
                {
                    _isIndexLoading = value;
                    OnPropertyChanged(nameof(IsIndexLoading));
                }
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether this instance is index updating.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if this instance is index updating; otherwise, <c>false</c>.
        /// </value>
        public bool IsIndexUpdating
        {
            get => _isIndexUpdating;
            set
            {
                if (_isIndexUpdating != value)
                {
                    _isIndexUpdating = value;
                    OnPropertyChanged(nameof(IsIndexUpdating));
                }
            }
        }

        /// <summary>
        /// Reads the snippet object and adds the right data to the index
        /// </summary>
        /// <param name="currentSnippet"></param>
        /// <param name="filePath">the path of the file</param>
        public void CreateIndexItemDataFromSnippet(Snippet currentSnippet, string filePath)
        {
            var item = new SnippetIndexItem();
            UpdateIndexItemData(item, currentSnippet);
            item.File = filePath;

            lock (_indexedSnippets)
            {
                _indexedSnippets[GetDictionaryKey(filePath, item.Title)] = item;
            }
        }

        /// <summary>
        /// Create a new index file by reading all snippet files
        /// and building them in internal memory then writing them to the index file
        /// </summary>
        public bool CreateOrUpdateIndexFile()
        {
            try
            {
                IsIndexUpdating = true;
                foreach (var path in SnippetDesignerPackage.Instance.Settings.AllSnippetDirectories)
                {
                    if (!Directory.Exists(path))
                    {
                        continue;
                    }

                    foreach (var snippetPath in Directory.GetFiles(path, SnippetSearch.AllSnippets, SearchOption.AllDirectories))
                    {
                        AddOrUpdateSnippetsToIndexFromSnippetFile(snippetPath);
                    }
                }

                IsIndexUpdating = false;

                //write the snippetitemcolllection to disk
                return SaveIndexFile();
            }
            catch (Exception e)
            {
                _logger.Log("Unable to read snippet index directories", "SnippetIndex", e);
            }

            return false;
        }

        /// <summary>
        /// delete the content associated with this item
        /// </summary>
        /// <param name="filePath">The file path.</param>
        /// <param name="title">The title.</param>
        public void DeleteSnippetFile(string filePath, string title)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                throw new ArgumentNullException(nameof(filePath), "filePath must not be null");
            }

            if (string.IsNullOrEmpty(title))
            {
                throw new ArgumentNullException(nameof(title), "title must not be null");
            }

            try
            {
                lock (_indexedSnippets)
                {
                    if (File.Exists(filePath))
                    {
                        //delete file on disk
                        File.Delete(filePath);
                    }

                    // Remove file from index
                    _indexedSnippets.Remove(GetDictionaryKey(filePath, title));

                    //save index file changes to disk
                    SaveIndexFile();
                }
            }
            catch (IOException e)
            {
                _logger.Log("Unable to delete snippet file", "SnippetIndex", e);
            }
        }

        /// <summary>
        /// Performs a search on the  snippets on the computer and return an index item collection containing them
        /// </summary>
        /// <param name="searchString">string to search by, if null or empty get all</param>
        /// <param name="languagesToGet">The languages to get.</param>
        /// <param name="maxResultCount">The max result count.</param>
        /// <returns>collection of found snippets</returns>
        public IEnumerable<SnippetIndexItem> PerformSnippetSearch(string searchString, List<string> languagesToGet, int maxResultCount)
        {
            var filterSnippets = _indexedSnippets.Where(x =>
                languagesToGet.Any(lang => lang.Equals(x.Value.Language, StringComparison.OrdinalIgnoreCase)));
            if (string.IsNullOrEmpty(searchString))
            {
                return filterSnippets.Select(x => x.Value).Take(maxResultCount);
            }

            var matchRankings = new List<Tuple<Regex, double>>
            {
                Tuple.Create(CreateRegex(@"\b{0}\b", searchString), 1.0), Tuple.Create(CreateRegex(@"{0}", searchString), 0.1)
            };

            var search = from snippet in filterSnippets
                         from fieldRanking in _fieldRankings
                         from matchRanking in matchRankings
                         where matchRanking.Item1.IsMatch(fieldRanking.Item1(snippet.Value))
                         let rank = fieldRanking.Item2 * matchRanking.Item2
                         group rank by snippet.Value
                into matchResults
                         orderby matchResults.Sum(x => x) descending
                         select matchResults.Key;

            return search.Take(maxResultCount);
        }

        /// <summary>
        /// Read the index file from disk into memory
        /// </summary>
        public bool ReadIndexFile()
        {
            IsIndexLoading = true;
            FileStream stream = null;
            try
            {
                //load the index file into memory
                stream = new FileStream(_snippetIndexFilePath, FileMode.Open);
                var items = Load(stream);
                if (items == null || items.Count == 0)
                {
                    return false;
                }

                foreach (var item in items)
                {
                    lock (_indexedSnippets)
                    {
                        if (File.Exists(item.File))
                        {
                            var key = GetDictionaryKey(item.File, item.Title);
                            if (!_indexedSnippets.ContainsKey(key))
                            {
                                _indexedSnippets.Add(key, item);
                            }
                        }
                    }
                }

                return true;
            }
            catch (Exception e)
            {
                _logger.Log("Unable to open snippet index file at path: " + _snippetIndexFilePath, "SnippetIndex", e);
                return false;
            }
            finally
            {
                stream?.Close();

                IsIndexLoading = false;
            }
        }

        /// <summary>
        /// Rebuilds the index of the snippet.
        /// </summary>
        public void RebuildSnippetIndex()
        {
            lock (_indexedSnippets)
            {
                _indexedSnippets.Clear();
            }

            CreateOrUpdateIndexFile();
        }

        /// <summary>
        /// Update a  snippet item in the collection based upon the current filepath
        /// then swap the item with the new one
        /// </summary>
        /// <param name="updatedSnippetFile"></param>
        public bool UpdateSnippetFile(SnippetFile updatedSnippetFile)
        {
            // Find keys to remove
            var keysToRemove = new List<string>();
            // Keys we found and updated
            var foundKeys = new List<string>();

            // These have title changes to we need to create a new key for them
            var snippetsToAdd = new List<Snippet>();

            // Update snippets that have not changed titles
            foreach (var snippet in updatedSnippetFile.Snippets)
            {
                var key = GetDictionaryKey(updatedSnippetFile.FileName, snippet.Title);
                _indexedSnippets.TryGetValue(key, out var item);
                if (item != null)
                {
                    UpdateIndexItemData(item, snippet);
                    foundKeys.Add(key);
                }
                else
                {
                    snippetsToAdd.Add(snippet);
                }
            }

            if (snippetsToAdd.Count > 0)
            {
                // Figure out which keys are no longer valid
                foreach (var key in _indexedSnippets.Keys)
                {
                    if (key.Contains(updatedSnippetFile.FileName.ToUpperInvariant()) &&
                        !foundKeys.Contains(key))
                    {
                        keysToRemove.Add(key);
                    }
                }

                // Since this file only has one snippet we know the one to update
                // so we don't need to re-add it
                if (updatedSnippetFile.Snippets.Count == 1 && keysToRemove.Count == 1)
                {
                    _indexedSnippets.TryGetValue(keysToRemove[0], out var item);
                    if (item != null)
                    {
                        UpdateIndexItemData(item, updatedSnippetFile.Snippets[0]);
                    }
                }
                else
                {
                    // Remove those keys
                    foreach (var key in keysToRemove)
                    {
                        lock (_indexedSnippets)
                        {
                            _indexedSnippets.Remove(key);
                        }
                    }

                    // Add update snippet items
                    foreach (var snippet in snippetsToAdd)
                    {
                        CreateIndexItemDataFromSnippet(snippet, updatedSnippetFile.FileName);
                    }
                }
            }

            return SaveIndexFile();
        }

        private static Regex CreateRegex(string pattern, string arg) =>
            new(string.Format(pattern, Regex.Escape(arg)), RegexOptions.IgnoreCase);

        /// <summary>
        /// Loads the data for this index item from a snippet file
        /// </summary>
        /// <param name="filePath">the path of the file</param>
        private bool AddOrUpdateSnippetsToIndexFromSnippetFile(string filePath)
        {
            try
            {
                var snippetFile = new SnippetFile(filePath);
                foreach (var currentSnippet in snippetFile.Snippets)
                {
                    _indexedSnippets.TryGetValue(GetDictionaryKey(filePath, currentSnippet.Title), out var existingItem);
                    if (existingItem == null)
                    {
                        //add the item to the collection
                        CreateIndexItemDataFromSnippet(currentSnippet, filePath);
                    }
                    else
                    {
                        UpdateIndexItemData(existingItem, currentSnippet);
                    }
                }
            }
            catch (IOException e)
            {
                _logger.Log("Unable to open snippet file at path: " + filePath, "SnippetIndex", e);
                return false;
            }

            return true;
        }

        /// <summary>
        /// Gets the dictionary key.
        /// </summary>
        /// <param name="filePath">The file path.</param>
        /// <param name="title">The title.</param>
        private static string GetDictionaryKey(string filePath, string title) =>
            filePath.ToUpperInvariant().Trim() + "|" + title.ToUpperInvariant().Trim();

        /// <summary>
        ///  Deserialize or Load this object member values from an XML file
        /// </summary>
        /// <param name="stream">Stream for the file to load</param>
        /// <returns>a List of snippetIndexItems or null if failure</returns>
        private List<SnippetIndexItem> Load(Stream stream)
        {
            if (stream == null)
            {
                return null;
            }

            var ser = new XmlSerializer(typeof(List<SnippetIndexItem>));
            return (List<SnippetIndexItem>)ser.Deserialize(stream);
        }

        /// <summary>
        /// Fire the property changed event for the given property name
        /// </summary>
        /// <param name="propertyName">Name of the property</param>
        private void OnPropertyChanged(string propertyName)
        {
            VerifyProperty(propertyName);
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        /// <summary>
        /// Serialize this object as an XML file to disk.
        /// </summary>
        /// <param name="stream">file stream</param>
        /// <returns>Succeed or failure</returns>
        private bool Save(Stream stream)
        {
            if (stream == null)
            {
                return false;
            }

            var ser = new XmlSerializer(typeof(List<SnippetIndexItem>));
            var items = new List<SnippetIndexItem>(_indexedSnippets.Values);
            ser.Serialize(stream, items);

            return true;
        }

        /// <summary>
        /// Write the current SnippetIndexItemCOllection to the index file
        /// </summary>
        /// <returns>true if success</returns>
        private bool SaveIndexFile()
        {
            //write the index to disk
            FileStream stream = null;
            try
            {
                var dirPath = Path.GetDirectoryName(_snippetIndexFilePath);
                Directory.CreateDirectory(dirPath);
                stream = new FileStream(_snippetIndexFilePath, FileMode.Create);
                return Save(stream);
            }
            catch (Exception e)
            {
                _logger.Log("Unable to write to snippet index file at path: " + _snippetIndexFilePath, "SnippetIndex", e);
            }
            finally
            {
                stream?.Close();
            }

            return false;
        }

        /// <summary>
        /// Update a snippet index item with new values
        /// </summary>
        /// <param name="item">item to update</param>
        /// <param name="snippetData">snippet data to update it with</param>
        private static void UpdateIndexItemData(SnippetIndexItem item, Snippet snippetData)
        {
            item.Title = snippetData.Title;
            item.Author = snippetData.Author;
            item.Description = snippetData.Description;
            item.Keywords = string.Join(",", snippetData.Keywords.ToArray());
            item.Language = snippetData.CodeLanguageAttribute;
            item.Code = snippetData.Code;
            item.Delimiter = snippetData.CodeDelimiterAttribute;
        }

        /// <summary>
        /// Used in debug mode only.  Checks to make sure the property name string
        /// is actually a property on the object.
        /// </summary>
        /// <param name="propertyName">Name of the property</param>
        [Conditional("DEBUG")]
        private void VerifyProperty(string propertyName)
        {
            var propertyExists = TypeDescriptor.GetProperties(this).Find(propertyName, false) != null;
            if (!propertyExists)
            {
                Debug.Fail(string.Format("The property {0} could not be found in {1}", propertyName, GetType().FullName));
            }
        }
    }
}
