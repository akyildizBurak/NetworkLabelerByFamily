using Autodesk.AutoCAD.Runtime;
using System.Windows;
using Autodesk.Civil.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.Civil.DatabaseServices;
using System;

namespace NetworkLabeler
{
    public class NetworkLabelerCommand
    {
        [CommandMethod("NETWORKLABELER")]
        public void ShowNetworkLabeler()
        {
            try
            {
                // Initialize logger
                DebugLogger.Initialize();
                DebugLogger.Log("Starting NetworkLabeler command");

                // Check Civil 3D document
                var civilDoc = CivilApplication.ActiveDocument;
                if (civilDoc == null)
                {
                    DebugLogger.LogError("No active Civil 3D document found");
                    MessageBox.Show("Please open a Civil 3D drawing first.", "NetworkLabeler Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Check if there are any pipe networks
                var networkIds = civilDoc.GetPipeNetworkIds();
                if (networkIds == null || networkIds.Count == 0)
                {
                    DebugLogger.LogError("No pipe networks found in the drawing");
                    MessageBox.Show("No pipe networks found in the current drawing.", "NetworkLabeler Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                DebugLogger.Log($"Found {networkIds.Count} networks. Creating main window...");
                var window = new MainWindow();
                window.ShowDialog();
            }
            catch (System.Exception ex)
            {
                DebugLogger.LogError("Error in NetworkLabeler command", ex);
                MessageBox.Show($"Error: {ex.Message}\nPlease check the log file for more details.", 
                    "NetworkLabeler Error", 
                    MessageBoxButton.OK, 
                    MessageBoxImage.Error);
            }
        }
    }
}