using System.Windows.Forms;
using System.Drawing;
using System;
using System.Runtime.InteropServices;
using SharpShell.Attributes;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Linq;

namespace ExtractMediaFilesExtension
{
    [ComVisible(true)]
    [COMServerAssociation(AssociationType.Directory)]
    public class ExtractMediaFilesExtension : SharpShell.SharpContextMenu.SharpContextMenu
    {
        protected override bool CanShowMenu()
        {
            if (SelectedItemPaths.Equals(string.Empty))
            {
                return false;
            }
            return true;
        }

        protected override ContextMenuStrip CreateMenu()
        {
            //  Create the menu strip.
            var menu = new ContextMenuStrip();

            //  Create a 'Extract media files' item.
            var itemExtractMediaFiles = new ToolStripMenuItem { Text = "Extract media files" };

            //  When we click, we'll call the 'ExtractMediaFiles' function.
            itemExtractMediaFiles.Click += (sender, args) => ExtractMediaFiles();

            //  Add the item to the context menu.
            menu.Items.Add(itemExtractMediaFiles);

            //  Return the menu.
            return menu;
        }

        private void ExtractMediaFiles()
        {
            try
            {
                List<string> extensions = new List<string> { "avi", "mp4", "mkv", "flv" };
                List<string> subFiles = new List<string>();
                string newFolderName = "Extracted Media Files";
                string startFolder = string.Empty;

                // Get all media subfiles
                foreach (var sip in SelectedItemPaths)
                {
                    startFolder = Directory.GetParent(sip).ToString(); // Does not need to run more than once, but who cares
                    foreach (string e in extensions)
                    {
                        string[] sf = Directory.GetFiles(sip, $"*.{e}");
                        if (sf.Length > 0)
                        {
                            subFiles.AddRange(sf);
                        }
                    }
                }

                // Figure out the new subfolder name
                foreach (var sf in subFiles)
                {
                    string fileName = Path.GetFileNameWithoutExtension(sf);
                    int indexOf = 0;
                    var match = Regex.Match(fileName, @"((\.[sS][0-9]{2}[eE][0-9]{2}))");
                    //MessageBox.Show($"match: {match.ToString()}");
                    indexOf = match.Index;
                    if (indexOf > 0)
                    {
                        newFolderName = fileName.Substring(0, indexOf);
                        break;
                    }
                }

                // Have the user approve the new foldername
                string value = newFolderName;
                if (MessageBoxes.InputBox("Extraction folder", "Folder name:", ref value) == DialogResult.OK)
                {
                    newFolderName = value;
                }
                else
                {
                    return;
                }

                // Create the new subfolder (if it does not exist)
                string newFolderPath = startFolder + @"\" + newFolderName;
                if (!Directory.Exists(newFolderPath)) Directory.CreateDirectory(newFolderPath);

                // Move the media files
                List<string> movedFiles = new List<string>();
                foreach (var f in subFiles)
                {
                    string newFile = newFolderPath + @"\" + Path.GetFileName(f);
                    if (!File.Exists(newFile))
                    {
                        try
                        {
                            File.Move(f, newFile);
                            movedFiles.Add(Path.GetFileName(f));
                        }
                        catch (Exception e)
                        {
                            MessageBoxes.OutputBox($"Moving violation: {Path.GetFileName(f)}", $"Error: {e.Message}");
                        }
                    }
                }

                // Go tell it on the mountain
                movedFiles.Sort();
                if (movedFiles.Count == 0) movedFiles.Add("No files were moved");
                string @out = movedFiles.Aggregate(new StringBuilder(),
                      (sb, a) => sb.AppendLine(String.Join(",", a)),
                      sb => sb.ToString());

                MessageBoxes.OutputBox("Moved files", @out);
            }
            catch (Exception e)
            {
                MessageBoxes.OutputBox("Exception", e.Message);
            }
        }
    }
}
