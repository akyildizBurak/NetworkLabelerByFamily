using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Linq;
using System.Runtime.InteropServices;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.Settings;
using Autodesk.Civil.DatabaseServices.Styles;
using AcadException = Autodesk.AutoCAD.Runtime.Exception;
using SystemException = System.Exception;
using AcadApplication = Autodesk.AutoCAD.ApplicationServices.Application;
using WpfApplication = System.Windows.Application;

namespace NetworkLabeler
{
    public class PartFamilyData : INotifyPropertyChanged
    {
        private string _familyName;
        private string _selectedStyle;
        private ObservableCollection<string> _availableLabelStyles;
        private bool _isPipe;
        private int _partCount;

        public string Name
        {
            get => _familyName;
            set
            {
                _familyName = value;
                OnPropertyChanged(nameof(Name));
            }
        }

        public int Count
        {
            get => _partCount;
            set
            {
                _partCount = value;
                OnPropertyChanged(nameof(Count));
            }
        }

        public string SelectedLabelStyle
        {
            get => _selectedStyle;
            set
            {
                _selectedStyle = value;
                OnPropertyChanged(nameof(SelectedLabelStyle));
            }
        }

        public ObservableCollection<string> LabelStyles
        {
            get => _availableLabelStyles;
            set
            {
                _availableLabelStyles = value;
                OnPropertyChanged(nameof(LabelStyles));
            }
        }

        public bool IsPipe
        {
            get => _isPipe;
            set
            {
                _isPipe = value;
                OnPropertyChanged(nameof(IsPipe));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        public virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        private readonly LabelStyleManager _labelStyleManager;
        private readonly CivilDocument _civilDoc;
        private readonly ObservableCollection<PartFamilyData> _partFamilies;

        public MainWindow()
        {
            try
            {
                Logger.Log("Starting MainWindow initialization");
                
                // Initialize collections first
                _partFamilies = new ObservableCollection<PartFamilyData>();
                Logger.Log("Part families collection initialized");

                // Initialize UI components
                InitializeComponent();
                if (PartFamiliesGrid != null)
                {
                    PartFamiliesGrid.ItemsSource = _partFamilies;
                    Logger.Log("PartFamiliesGrid ItemsSource set");
                }
                Logger.Log("MainWindow UI components initialized");
                
                // Get the active Civil 3D document
                var acDoc = AcadApplication.DocumentManager.MdiActiveDocument;
                if (acDoc == null)
                {
                    throw new SystemException("No active AutoCAD document found. Please open a drawing first.");
                }
                Logger.Log("Active AutoCAD document found");

                _civilDoc = CivilApplication.ActiveDocument;
                if (_civilDoc == null)
                {
                    throw new SystemException("No active Civil 3D document found. Please open a Civil 3D drawing.");
                }
                Logger.Log("Active Civil 3D document found");

                // Initialize label style manager
                _labelStyleManager = new LabelStyleManager(_civilDoc);
                if (_labelStyleManager == null)
                {
                    throw new SystemException("Failed to initialize label style manager.");
                }
                Logger.Log("Label style manager initialized");
                
                // Load saved label style selections
                if (acDoc.Name != null)
                {
                    _labelStyleManager.LoadSelections(acDoc.Name);
                    Logger.Log("Loaded saved label style selections");
                }

                // Load available networks
                LoadNetworks();
                Logger.Log("MainWindow initialization completed successfully");
            }
            catch (AcadException ex)
            {
                Logger.LogError("AutoCAD error in MainWindow initialization", ex);
                MessageBox.Show($"AutoCAD error: {ex.Message}\nPlease ensure you have a Civil 3D drawing open.", 
                    "Initialization Error", 
                    MessageBoxButton.OK, 
                    MessageBoxImage.Error);
                Close();
            }
            catch (SystemException ex)
            {
                Logger.LogError("Error initializing MainWindow", ex);
                MessageBox.Show($"Error initializing window: {ex.Message}\nPlease ensure you have a Civil 3D drawing open.", 
                    "Initialization Error", 
                    MessageBoxButton.OK, 
                    MessageBoxImage.Error);
                Close(); // Close the window if initialization fails
            }
        }

        private void LoadNetworks()
        {
            try
            {
                if (_civilDoc == null)
                {
                    throw new SystemException("No active Civil 3D document.");
                }

                if (_labelStyleManager == null)
                {
                    throw new SystemException("Label style manager is not initialized.");
                }

                var networkIds = _civilDoc.GetPipeNetworkIds();
                if (networkIds == null || networkIds.Count == 0)
                {
                    MessageBox.Show("No pipe networks found in the current drawing.", 
                        "No Networks", 
                        MessageBoxButton.OK, 
                        MessageBoxImage.Information);
                    return;
                }

                var networks = new List<string>();
                var db = networkIds[0].Database;
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    foreach (ObjectId networkId in networkIds)
                    {
                        if (networkId.IsNull || networkId.IsErased)
                            continue;

                        var network = tr.GetObject(networkId, OpenMode.ForRead) as Network;
                        if (network != null && !string.IsNullOrEmpty(network.Name))
                        {
                            networks.Add(network.Name);
                        }
                    }
                    tr.Commit();
                }

                if (networks.Count == 0)
                {
                    MessageBox.Show("No valid pipe networks found in the current drawing.", 
                        "No Networks", 
                        MessageBoxButton.OK, 
                        MessageBoxImage.Information);
                    return;
                }

                if (NetworkSelect != null)
                {
                    NetworkSelect.ItemsSource = networks;
                    NetworkSelect.SelectedIndex = 0;
                }
                
                if (networks.Count > 0)
                {
                    var selectedNetwork = GetNetworkByName(networks[0]);
                    if (!selectedNetwork.IsNull)
                    {
                        LoadPartFamilies(selectedNetwork);
                    }
                }
            }
            catch (SystemException ex)
            {
                MessageBox.Show($"Error loading networks: {ex.Message}\nStack trace: {ex.StackTrace}", 
                    "Error", 
                    MessageBoxButton.OK, 
                    MessageBoxImage.Error);
            }
        }

        private ObjectId GetNetworkByName(string networkName)
        {
            try
            {
                var networkIds = _civilDoc.GetPipeNetworkIds();
                if (networkIds == null || networkIds.Count == 0)
                    return ObjectId.Null;

                var db = networkIds[0].Database;
                ObjectId resultId = ObjectId.Null;

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    foreach (ObjectId networkId in networkIds)
                    {
                        if (!networkId.IsNull && !networkId.IsErased)
                        {
                            var network = tr.GetObject(networkId, OpenMode.ForRead) as Network;
                            if (network != null && network.Name == networkName)
                            {
                                resultId = networkId;
                                break;
                            }
                        }
                    }
                    tr.Commit();
                }
                return resultId;
            }
            catch (SystemException ex)
            {
                MessageBox.Show($"Error getting network: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return ObjectId.Null;
            }
        }

        private void NetworkSelect_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (NetworkSelect.SelectedItem is string networkName)
            {
                var networkId = GetNetworkByName(networkName);
                if (!networkId.IsNull)
                {
                    LoadPartFamilies(networkId);
                }
            }
        }

        private void LoadPartFamilies(ObjectId networkId)
        {
            Transaction tr = null;
            try
            {
                Logger.Log("=== Starting LoadPartFamilies ===");
                
                if (networkId.IsNull)
                {
                    Logger.Log("ERROR: Network ObjectId is null");
                    MessageBox.Show("Network ID is null", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                Logger.Log($"Network ObjectId is valid: {networkId}");

                if (_labelStyleManager == null)
                {
                    Logger.Log("ERROR: Label style manager is null");
                    MessageBox.Show("Label style manager is not initialized", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                Logger.Log("Label style manager is valid");

                if (_partFamilies == null)
                {
                    Logger.Log("ERROR: Part families collection is null");
                    MessageBox.Show("Part families collection is not initialized", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                Logger.Log("Part families collection is valid");

                var db = networkId.Database;
                if (db == null)
                {
                    Logger.Log("ERROR: Database is null");
                    MessageBox.Show("Database is null", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                Logger.Log("Database is valid");

                // Get all available label styles first
                Logger.Log("Getting pipe label styles...");
                var pipeStyles = _labelStyleManager.GetPipeLabelStyles();
                Logger.Log($"Retrieved {pipeStyles.Count} pipe styles");
                
                Logger.Log("Getting structure label styles...");
                var structureStyles = _labelStyleManager.GetStructureLabelStyles();
                Logger.Log($"Retrieved {structureStyles.Count} structure styles");

                // Clear existing data
                Logger.Log("Clearing existing part families...");
                try
                {
                    if (WpfApplication.Current != null)
                    {
                        WpfApplication.Current.Dispatcher.Invoke(() =>
                        {
                            _partFamilies.Clear();
                        });
                        Logger.Log("Cleared existing part families (via dispatcher)");
                    }
                    else
                    {
                        _partFamilies.Clear();
                        Logger.Log("Cleared existing part families (direct)");
                    }
                }
                catch (SystemException ex)
                {
                    Logger.LogError("Error clearing part families", ex);
                    throw;
                }

                var partsByFamily = new Dictionary<string, PartFamilyData>();

                Logger.Log("Starting transaction...");
                tr = db.TransactionManager.StartTransaction();
                Logger.Log("Transaction started successfully");

                // Get the network object within this transaction
                Logger.Log($"Attempting to get network object for ID: {networkId}");
                var network = tr.GetObject(networkId, OpenMode.ForRead) as Network;
                if (network == null)
                {
                    Logger.Log("ERROR: Failed to access network object");
                    throw new SystemException("Failed to access network object.");
                }

                Logger.Log($"Successfully accessed network: {network.Name}");

                // Get pipes
                Logger.Log("Getting pipe IDs...");
                var pipes = network.GetPipeIds();
                Logger.Log($"Found {(pipes?.Count ?? 0)} pipes");
                
                if (pipes != null)
                {
                    foreach (ObjectId pipeId in pipes)
                    {
                        if (!pipeId.IsNull && !pipeId.IsErased)
                        {
                            Logger.Log($"Processing pipe ID: {pipeId}");
                            var pipe = tr.GetObject(pipeId, OpenMode.ForRead) as Pipe;
                            if (pipe != null)
                            {
                                var familyName = pipe.PartFamilyName;
                                Logger.Log($"Found pipe family: {familyName}");
                                if (!string.IsNullOrEmpty(familyName))
                                {
                                    if (!partsByFamily.ContainsKey(familyName))
                                    {
                                        var savedStyle = _labelStyleManager.GetSelectedLabelStyle(familyName);
                                        Logger.Log($"Retrieved saved style for {familyName}: {savedStyle ?? "none"}");
                                        
                                        var styles = new ObservableCollection<string>(pipeStyles);
                                        var selectedStyle = savedStyle;
                                        if (string.IsNullOrEmpty(selectedStyle) || !styles.Contains(selectedStyle))
                                        {
                                            selectedStyle = styles.FirstOrDefault();
                                        }

                                        var newFamily = new PartFamilyData
                                        {
                                            Name = familyName,
                                            LabelStyles = styles,
                                            SelectedLabelStyle = selectedStyle,
                                            IsPipe = true,
                                            Count = 1
                                        };

                                        partsByFamily[familyName] = newFamily;
                                        AddToPartFamilies(newFamily);
                                    }
                                    else
                                    {
                                        partsByFamily[familyName].Count++;
                                    }
                                }
                            }
                        }
                    }
                }

                // Get structures
                Logger.Log("Getting structure IDs...");
                var structures = network.GetStructureIds();
                Logger.Log($"Found {(structures?.Count ?? 0)} structures");
                
                if (structures != null)
                {
                    foreach (ObjectId structureId in structures)
                    {
                        if (!structureId.IsNull && !structureId.IsErased)
                        {
                            Logger.Log($"Processing structure ID: {structureId}");
                            var structure = tr.GetObject(structureId, OpenMode.ForRead) as Structure;
                            if (structure != null)
                            {
                                var familyName = structure.PartFamilyName;
                                Logger.Log($"Found structure family: {familyName}");
                                if (!string.IsNullOrEmpty(familyName))
                                {
                                    if (!partsByFamily.ContainsKey(familyName))
                                    {
                                        var savedStyle = _labelStyleManager.GetSelectedLabelStyle(familyName);
                                        Logger.Log($"Retrieved saved style for {familyName}: {savedStyle ?? "none"}");
                                        
                                        var styles = new ObservableCollection<string>(structureStyles);
                                        var selectedStyle = savedStyle;
                                        if (string.IsNullOrEmpty(selectedStyle) || !styles.Contains(selectedStyle))
                                        {
                                            selectedStyle = styles.FirstOrDefault();
                                        }

                                        var newFamily = new PartFamilyData
                                        {
                                            Name = familyName,
                                            LabelStyles = styles,
                                            SelectedLabelStyle = selectedStyle,
                                            IsPipe = false,
                                            Count = 1
                                        };

                                        partsByFamily[familyName] = newFamily;
                                        AddToPartFamilies(newFamily);
                                    }
                                    else
                                    {
                                        partsByFamily[familyName].Count++;
                                    }
                                }
                            }
                        }
                    }
                }

                tr.Commit();
                Logger.Log("=== LoadPartFamilies completed successfully ===");
            }
            catch (SystemException ex)
            {
                Logger.LogError("Error in LoadPartFamilies", ex);
                tr?.Abort();
                MessageBox.Show($"Error loading part families: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void AddToPartFamilies(PartFamilyData newFamily)
        {
            try
            {
                if (WpfApplication.Current != null)
                {
                    WpfApplication.Current.Dispatcher.Invoke(() =>
                    {
                        _partFamilies.Add(newFamily);
                    });
                    Logger.Log($"Added family: {newFamily.Name} (via dispatcher)");
                }
                else
                {
                    _partFamilies.Add(newFamily);
                    Logger.Log($"Added family: {newFamily.Name} (direct)");
                }
            }
            catch (SystemException ex)
            {
                Logger.LogError($"Error adding family {newFamily.Name}", ex);
                // Try direct add as fallback
                try
                {
                    _partFamilies.Add(newFamily);
                    Logger.Log($"Added family: {newFamily.Name} (fallback)");
                }
                catch (SystemException addEx)
                {
                    Logger.LogError($"Failed to add family {newFamily.Name}", addEx);
                    throw; // Re-throw if we really can't add to the collection
                }
            }
        }

        private void ApplyStyle_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                if (doc == null)
                {
                    MessageBox.Show("No active document found.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                var db = doc.Database;
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    try
                    {
                        Logger.Log("Starting to apply label styles...");

                        // Save selected styles before applying
                        foreach (var family in _partFamilies)
                        {
                            if (!string.IsNullOrEmpty(family.SelectedLabelStyle))
                            {
                                _labelStyleManager.SetSelectedLabelStyle(family.Name, family.SelectedLabelStyle);
                                Logger.Log($"Saved style {family.SelectedLabelStyle} for family {family.Name}");
                            }
                        }
                        
                        // Save to file
                        if (doc.Name != null)
                        {
                            _labelStyleManager.SaveSelections(doc.Name);
                            Logger.Log("Saved label style selections to file");
                        }

                        // Get the current network
                        if (NetworkSelect.SelectedItem == null)
                        {
                            throw new SystemException("No network selected.");
                        }

                        var networkId = GetNetworkByName(NetworkSelect.SelectedItem.ToString());
                        if (networkId.IsNull)
                        {
                            throw new SystemException("Selected network not found.");
                        }

                        var network = tr.GetObject(networkId, OpenMode.ForRead) as Network;
                        if (network == null)
                        {
                            throw new SystemException("Failed to access network object.");
                        }

                        Logger.Log($"Processing network: {network.Name}");
                        int labelCount = 0;

                        // Process pipes
                        var pipeIds = network.GetPipeIds();
                        if (pipeIds != null)
                        {
                            foreach (ObjectId pipeId in pipeIds)
                            {
                                if (!pipeId.IsNull && !pipeId.IsErased)
                                {
                                    var pipe = tr.GetObject(pipeId, OpenMode.ForWrite) as Pipe;
                                    if (pipe != null)
                                    {
                                        var familyData = _partFamilies.FirstOrDefault(pf => pf.Name == pipe.PartFamilyName && pf.IsPipe);
                                        if (familyData != null && !string.IsNullOrEmpty(familyData.SelectedLabelStyle))
                                        {
                                            _labelStyleManager.ApplyLabelStyle(pipe, familyData.SelectedLabelStyle);
                                            labelCount++;
                                            Logger.Log($"Applied style {familyData.SelectedLabelStyle} to pipe of family {familyData.Name}");
                                        }
                                    }
                                }
                            }
                        }

                        // Process structures
                        var structureIds = network.GetStructureIds();
                        if (structureIds != null)
                        {
                            foreach (ObjectId structureId in structureIds)
                            {
                                if (!structureId.IsNull && !structureId.IsErased)
                                {
                                    var structure = tr.GetObject(structureId, OpenMode.ForWrite) as Structure;
                                    if (structure != null)
                                    {
                                        var familyData = _partFamilies.FirstOrDefault(pf => pf.Name == structure.PartFamilyName && !pf.IsPipe);
                                        if (familyData != null && !string.IsNullOrEmpty(familyData.SelectedLabelStyle))
                                        {
                                            _labelStyleManager.ApplyLabelStyle(structure, familyData.SelectedLabelStyle);
                                            labelCount++;
                                            Logger.Log($"Applied style {familyData.SelectedLabelStyle} to structure of family {familyData.Name}");
                                        }
                                    }
                                }
                            }
                        }

                        tr.Commit();
                        Logger.Log($"Successfully applied {labelCount} labels");
                        MessageBox.Show($"Successfully applied label styles to {labelCount} parts!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    catch (SystemException ex)
                    {
                        Logger.LogError("Error while applying label styles", ex);
                        tr.Abort();
                        MessageBox.Show($"Error applying label styles: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
            catch (SystemException ex)
            {
                Logger.LogError("Error in ApplyStyle_Click", ex);
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SelectByPick_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                var ed = doc.Editor;
                var db = doc.Database;
                var tr = db.TransactionManager.StartTransaction();

                try
                {
                    // Prompt user to select an object
                    var pso = new PromptSelectionOptions();
                    pso.MessageForAdding = "\nSelect a pipe or structure: ";
                    var res = ed.GetSelection(pso);

                    if (res.Status == PromptStatus.OK)
                    {
                        foreach (SelectedObject selObj in res.Value)
                        {
                            var ent = tr.GetObject(selObj.ObjectId, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Entity;
                            if (ent != null)
                            {
                                if (ent is Pipe pipe)
                                {
                                    var networkIds = _civilDoc.GetPipeNetworkIds();
                                    foreach (ObjectId networkId in networkIds)
                                    {
                                        var network = tr.GetObject(networkId, OpenMode.ForRead) as Network;
                                        if (network != null)
                                        {
                                            var pipeIds = network.GetPipeIds();
                                            if (pipeIds != null && pipeIds.Contains(pipe.ObjectId))
                                            {
                                                NetworkSelect.Text = network.Name;
                                                LoadPartFamilies(networkId);
                                                break;
                                            }
                                        }
                                    }
                                    break;
                                }
                                else if (ent is Structure structure)
                                {
                                    var networkIds = _civilDoc.GetPipeNetworkIds();
                                    foreach (ObjectId networkId in networkIds)
                                    {
                                        var network = tr.GetObject(networkId, OpenMode.ForRead) as Network;
                                        if (network != null)
                                        {
                                            var structureIds = network.GetStructureIds();
                                            if (structureIds != null && structureIds.Contains(structure.ObjectId))
                                            {
                                                NetworkSelect.Text = network.Name;
                                                LoadPartFamilies(networkId);
                                                break;
                                            }
                                        }
                                    }
                                    break;
                                }
                            }
                        }
                    }
                    tr.Commit();
                }
                catch (SystemException ex)
                {
                    tr.Abort();
                    MessageBox.Show($"Error during selection: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (SystemException ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void AddFamily_Click(object sender, RoutedEventArgs e)
        {
            // Implementation for adding a family
        }

        private void RemoveFamily_Click(object sender, RoutedEventArgs e)
        {
            // Implementation for removing a family
        }

        private void SaveConfig_Click(object sender, RoutedEventArgs e)
        {
            // Implementation for saving configuration
        }

        private void LoadConfig_Click(object sender, RoutedEventArgs e)
        {
            // Implementation for loading configuration
        }
    }
} 