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

            //  Create a 'count lines' item.
            var itemCountLines = new ToolStripMenuItem
            {
                Text = "Extract media files"
            };

            //  When we click, we'll call the 'CountLines' function.
            itemCountLines.Click += (sender, args) => ExtractMediaFiles();

            //  Add the item to the context menu.
            menu.Items.Add(itemCountLines);

            //  Return the menu.
            return menu;
        }

        private void ExtractMediaFiles()
        {
            List<string> extentions = new List<string> { "avi", "mp4", "mkv" };
            List<string> subFiles = new List<string>();
            string newFolderName = "Extracted Media Files";
            string startFolder = string.Empty;

            // Get all media subfiles
            foreach (var sip in SelectedItemPaths)
            {
                startFolder = Directory.GetParent(sip).ToString();
                foreach (string e in extentions)
                {
                    string[] sf = Directory.GetFiles(sip, $"*.{e}");
                    if (sf.Length > 0)
                    {
                        subFiles.AddRange(sf);
                    }
                }
            }

            // Figure out the subfolder name
            foreach (var sf in subFiles)
            {
                string fileName = Path.GetFileNameWithoutExtension(sf);
                int indexOf = 0;
                //MessageBox.Show($": {}");
                var match = Regex.Match(fileName, @"((\.[sS][0-9]{2}[eE][0-9]{2}))");
                indexOf = match.Index;
                if (indexOf > 0)
                {
                    newFolderName = fileName.Substring(0, indexOf);
                    break;
                }
            }

            // Have the user approve the new foldername
            string value = newFolderName;
            if (InputBox("Extraction folder", "Folder name:", ref value) == DialogResult.OK)
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
                    File.Move(f, newFile);
                }
                movedFiles.Add(Path.GetFileName(f));
            }

            movedFiles.Sort();
            string @out = movedFiles.Aggregate(new StringBuilder(),
                  (sb, a) => sb.AppendLine(String.Join(",", a)),
                  sb => sb.ToString());
            MessageBox.Show(@out);

        }

        public static DialogResult InputBox(string title, string promptText, ref string value)
        {
            Form form = new Form();
            Label label = new Label();
            TextBox textBox = new TextBox();
            Button buttonOk = new Button();
            Button buttonCancel = new Button();

            form.Text = title;
            label.Text = promptText;
            textBox.Text = value;

            buttonOk.Text = "OK";
            buttonCancel.Text = "Cancel";
            buttonOk.DialogResult = DialogResult.OK;
            buttonCancel.DialogResult = DialogResult.Cancel;

            label.SetBounds(9, 20, 372, 13);
            textBox.SetBounds(12, 36, 372, 20);
            buttonOk.SetBounds(228, 72, 75, 23);
            buttonCancel.SetBounds(309, 72, 75, 23);

            label.AutoSize = true;
            textBox.Anchor = textBox.Anchor | AnchorStyles.Right;
            buttonOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            buttonCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;

            form.ClientSize = new Size(396, 107);
            form.Controls.AddRange(new Control[] { label, textBox, buttonOk, buttonCancel });
            form.ClientSize = new Size(Math.Max(300, label.Right + 10), form.ClientSize.Height);
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.StartPosition = FormStartPosition.CenterScreen;
            form.MinimizeBox = false;
            form.MaximizeBox = false;
            form.AcceptButton = buttonOk;
            form.CancelButton = buttonCancel;

            DialogResult dialogResult = form.ShowDialog();
            value = textBox.Text;
            return dialogResult;
        }
    }
}
