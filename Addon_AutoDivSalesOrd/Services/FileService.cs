using Addon_AutoDivSalesOrd.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Addon_AutoDivSalesOrd.Services
{
    /// <summary>
    /// Service responsible for file-related operations.
    /// </summary>
    public class FileService
    {
        /// <summary>
        /// Opens a file dialog for selecting a .txt file
        /// and returns the full path of the selected file.
        /// </summary>
        /// <returns>Full file path or null if cancelled.</returns>
        public static string GetTxtFilePath()
        {
            try
            {
                string result = null;

                var thread = new Thread(() =>
                {
                    using (var dialog = new OpenFileDialog
                    {
                        Filter = "Text Files (*.txt)|*.txt",
                        Title = "Select a .txt file",
                        CheckFileExists = true,
                        Multiselect = false
                    })
                    {
                        var dlgResult = dialog.ShowDialog();
                        if (dlgResult == DialogResult.OK)
                            result = dialog.FileName;
                    }
                });

                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
                thread.Join();

                return result;
            }
            catch (Exception ex)
            {
                throw new Exception(Constants.Messages.TxtFilePathErrorPrefix + ex.Message);
            }
        }
    }
}
