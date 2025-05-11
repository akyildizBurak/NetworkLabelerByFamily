using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
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
        private string _currentStyle;
        private ObservableCollection<string> _availableLabelStyles;
        private bool _isPipe;
        private int _partCount;
        private HashSet<string> _currentStyles = new HashSet<string>();
        private bool _isSelected;
        private List<Part> _parts = new List<Part>();
        private string _currentSelectedStyle = "No Label";
        private ObservableCollection<string> _currentStylesList = new ObservableCollection<string> { "No Label" };

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

        private ObservableCollection<string> _labelStyles;
        public ObservableCollection<string> LabelStyles 
        { 
            get 
            {
                if (_labelStyles == null)
                {
                    _labelStyles = new ObservableCollection<string> { "Select New Style..." };
                }
                
                // Add "Delete Style" option for multiple styles
                if (_labelStyles.Count > 1)
                {
                    if (!_labelStyles.Contains("Delete Style"))
                    {
                        _labelStyles.Add("Delete Style");
                    }
                }
                
                return _labelStyles;
            }
            set
            {
                _labelStyles = value;
                OnPropertyChanged(nameof(LabelStyles));
            }
        }

        private string _selectedLabelStyle = "Select New Style...";
        public string SelectedLabelStyle
        {
            get => _selectedLabelStyle;
            set
            {
                if (_selectedLabelStyle != value)
                {
                    _selectedLabelStyle = value;
                    OnPropertyChanged(nameof(SelectedLabelStyle));
                }
            }
        }

        public string CurrentStyle
        {
            get => _currentStyle;
            set
            {
                _currentStyle = value;
                OnPropertyChanged(nameof(CurrentStyle));
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

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                _isSelected = value;
                OnPropertyChanged(nameof(IsSelected));
            }
        }

        public List<Part> Parts
        {
            get => _parts;
            set
            {
                _parts = value;
                OnPropertyChanged(nameof(Parts));
            }
        }

        public HashSet<string> CurrentStyles
        {
            get => _currentStyles;
            set
            {
                _currentStyles = value ?? new HashSet<string>();
                UpdateCurrentStyleString();
                OnPropertyChanged(nameof(CurrentStyles));
            }
        }

        public ObservableCollection<string> CurrentStylesList
        {
            get => _currentStylesList;
            set
            {
                _currentStylesList = value;
                OnPropertyChanged(nameof(CurrentStylesList));
            }
        }

        public string CurrentSelectedStyle
        {
            get => _currentSelectedStyle;
            set
            {
                _currentSelectedStyle = value;
                OnPropertyChanged(nameof(CurrentSelectedStyle));
            }
        }

        public void UpdateCurrentStyleString()
        {
            var styles = _currentStyles.Where(s => !string.IsNullOrEmpty(s)).OrderBy(s => s).ToList();
            
            // Update CurrentStyle string
            CurrentStyle = styles.Count > 0 
                ? string.Join(", ", styles)
                : "No Label";
            
            // Update CurrentStylesList
            CurrentStylesList.Clear();
            
            if (styles.Count == 0)
            {
                // No styles at all
                CurrentStylesList.Add("No Label");
                CurrentSelectedStyle = "No Label";
            }
            else if (styles.Count == 1)
            {
                // Only one style, preselect it
                CurrentStylesList.Add(styles[0]);
                CurrentSelectedStyle = styles[0];
            }
            else
            {
                // Multiple styles
                CurrentStylesList.Add("Select Style...");
                CurrentStylesList.Add("All Styles");
                
                // Add individual styles
                foreach (var style in styles)
                {
                    CurrentStylesList.Add(style);
                }
                
                // Default to "Select Style..."
                CurrentSelectedStyle = "Select Style...";
            }

            OnPropertyChanged(nameof(CurrentStylesList));
            OnPropertyChanged(nameof(CurrentSelectedStyle));
        }

        public void AddCurrentStyle(string style)
        {
            if (!string.IsNullOrEmpty(style))
            {
                // Split comma-separated styles and trim whitespace
                var stylesToAdd = style
                    .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(s => s.Trim())
                    .Where(s => !string.IsNullOrEmpty(s));

                bool stylesAdded = false;
                foreach (var individualStyle in stylesToAdd)
                {
                    // Ensure we don't add duplicate styles
                    if (!_currentStyles.Contains(individualStyle))
                    {
                        _currentStyles.Add(individualStyle);
                        stylesAdded = true;
                    }
                }

                // Update only if new styles were added
                if (stylesAdded)
                {
                    UpdateCurrentStyleString();
                }
            }
        }

        public void AddPart(Part part)
        {
            if (part != null && !_parts.Contains(part))
            {
                _parts.Add(part);
                Count = _parts.Count;
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
    public partial class MainWindow : Window
    {
        private readonly LabelStyleManager _labelStyleManager;
        private readonly CivilDocument _civilDoc;
        private readonly ObservableCollection<PartFamilyData> _partFamilies;
        private readonly Document _acadDoc;
        private IntPtr _acadWindowHandle;

        public MainWindow()
        {
            try
            {
                Logger.Log("Starting MainWindow initialization");
                
                // Store the active document reference first
                _acadDoc = AcadApplication.DocumentManager.MdiActiveDocument;
                if (_acadDoc == null)
                {
                    throw new SystemException("No active AutoCAD document found. Please open a drawing first.");
                }
                Logger.Log("Active AutoCAD document found");

                // Get AutoCAD window handle
                _acadWindowHandle = _acadDoc.Window.Handle;

                // Set window ownership before initializing components
                WindowInteropHelper helper = new WindowInteropHelper(this);
                helper.Owner = _acadWindowHandle;

                // Initialize collections
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
                if (_acadDoc.Name != null)
                {
                    _labelStyleManager.LoadSelections(_acadDoc.Name);
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
                Close();
            }
        }

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            // Make sure window stays visible but doesn't block AutoCAD
            this.Topmost = false;
            this.ShowInTaskbar = true;
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
                    // Add a default "Select a Network" item
                    networks.Insert(0, "Select a Network");
                    NetworkSelect.SelectedIndex = 0;
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
            // Clear part families if no network is selected
            if (_partFamilies != null)
            {
                _partFamilies.Clear();
            }

            // Only proceed if a valid network is selected (not the default "Select a Network")
            if (NetworkSelect.SelectedItem is string networkName && 
                !string.IsNullOrWhiteSpace(networkName) && 
                networkName != "Select a Network")
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
                    // Try direct add as fallback
                    try
                    {
                        _partFamilies.Clear();
                        Logger.Log("Cleared existing part families (fallback)");
                    }
                    catch (SystemException addEx)
                    {
                        Logger.LogError($"Failed to clear part families", addEx);
                        throw; // Re-throw if we really can't clear the collection
                    }
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
                var pipeIds = network.GetPipeIds();
                Logger.Log($"Found {(pipeIds?.Count ?? 0)} pipes");
                
                if (pipeIds != null)
                {
                    foreach (ObjectId pipeId in pipeIds)
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
                                        styles.Insert(0, "Select New Style...");
                                        var selectedStyle = savedStyle;
                                        if (string.IsNullOrEmpty(selectedStyle) || !styles.Contains(selectedStyle))
                                        {
                                            selectedStyle = styles.FirstOrDefault();
                                        }

                                        // Get current style from the pipe
                                        string currentStyle = _labelStyleManager.GetCurrentLabelStyle(pipe);
                                        Logger.Log($"Current style for pipe family {familyName}: {currentStyle ?? "none"}");

                                        var newFamily = new PartFamilyData
                                        {
                                            Name = familyName,
                                            LabelStyles = styles,
                                            SelectedLabelStyle = selectedStyle,
                                            CurrentStyle = currentStyle ?? "No Label",
                                            CurrentStyles = new HashSet<string>(),  // Initialize empty set
                                            IsPipe = true,
                                            Count = 1,
                                            IsSelected = true // Set selected by default
                                        };

                                        if (!string.IsNullOrEmpty(currentStyle))
                                        {
                                            newFamily.AddCurrentStyle(currentStyle);
                                        }

                                        newFamily.AddPart(pipe);
                                        partsByFamily[familyName] = newFamily;
                                        AddToPartFamilies(newFamily);
                                    }
                                    else
                                    {
                                        partsByFamily[familyName].Count++;
                                        // Update current styles
                                        string currentStyle = _labelStyleManager.GetCurrentLabelStyle(pipe);
                                        if (!string.IsNullOrEmpty(currentStyle))
                                        {
                                            partsByFamily[familyName].AddCurrentStyle(currentStyle);
                                        }
                                        partsByFamily[familyName].AddPart(pipe);
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
                                        styles.Insert(0, "Select New Style...");
                                        var selectedStyle = savedStyle;
                                        if (string.IsNullOrEmpty(selectedStyle) || !styles.Contains(selectedStyle))
                                        {
                                            selectedStyle = styles.FirstOrDefault();
                                        }

                                        // Get current style from the structure
                                        string currentStyle = _labelStyleManager.GetCurrentLabelStyle(structure);
                                        Logger.Log($"Current style for structure family {familyName}: {currentStyle ?? "none"}");

                                        var newFamily = new PartFamilyData
                                        {
                                            Name = familyName,
                                            LabelStyles = styles,
                                            SelectedLabelStyle = selectedStyle,
                                            CurrentStyle = currentStyle ?? "No Label",
                                            CurrentStyles = new HashSet<string>(),  // Initialize empty set
                                            IsPipe = false,
                                            Count = 1,
                                            IsSelected = true // Set selected by default
                                        };

                                        if (!string.IsNullOrEmpty(currentStyle))
                                        {
                                            newFamily.AddCurrentStyle(currentStyle);
                                        }

                                        newFamily.AddPart(structure);
                                        partsByFamily[familyName] = newFamily;
                                        AddToPartFamilies(newFamily);
                                    }
                                    else
                                    {
                                        partsByFamily[familyName].Count++;
                                        // Update current styles
                                        string currentStyle = _labelStyleManager.GetCurrentLabelStyle(structure);
                                        if (!string.IsNullOrEmpty(currentStyle))
                                        {
                                            partsByFamily[familyName].AddCurrentStyle(currentStyle);
                                        }
                                        partsByFamily[familyName].AddPart(structure);
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
                Logger.Log("Starting ApplyStyle_Click...");
                var selectedNetwork = NetworkSelect.SelectedItem as string;
                if (selectedNetwork == null)
                {
                    MessageBox.Show("Please select a network first.");
                    return;
                }

                bool replaceExisting = RbSwapExisting.IsChecked ?? true;
                int updatedCount = 0;
                int totalSelected = 0;

                // Use AutoCAD's document lock
                using (DocumentLock docLock = _acadDoc.LockDocument(DocumentLockMode.Write, "Applying Label Styles", "Updating labels...", false))
                {
                    foreach (var familyData in _partFamilies)
                    {
                        if (!familyData.IsSelected) continue;
                        totalSelected++;

                        // Skip if no label style is selected or is the placeholder
                        if (string.IsNullOrEmpty(familyData.SelectedLabelStyle) || 
                            familyData.SelectedLabelStyle == "Select New Style...")
                        {
                            Logger.Log($"No label style selected for family {familyData.Name} - skipping");
                            continue;
                        }

                        // Handle "Delete Style" option
                        if (familyData.SelectedLabelStyle == "Delete Style")
                        {
                            // Ensure a current style is selected for deletion
                            if (string.IsNullOrEmpty(familyData.CurrentSelectedStyle) || 
                                familyData.CurrentSelectedStyle == "Select Style..." ||
                                familyData.CurrentSelectedStyle == "No Label")
                            {
                                Logger.Log($"No current style selected for deletion in family {familyData.Name} - skipping");
                                continue;
                            }

                            // Delete the selected current style from parts
                            foreach (var part in familyData.Parts)
                            {
                                // Delete labels with the selected current style
                                _labelStyleManager.DeleteSpecificLabelStyle(part, familyData.CurrentSelectedStyle);
                                updatedCount++;
                            }

                            // Remove the deleted style from current styles
                            familyData.CurrentStyles.Remove(familyData.CurrentSelectedStyle);

                            // Reset current style selection
                            familyData.CurrentSelectedStyle = "Select Style...";

                            // Update the current style string to reflect changes
                            familyData.UpdateCurrentStyleString();

                            continue;
                        }

                        // For "Add New Labels", process ALL parts in the selected family
                        if (!replaceExisting)
                        {
                            foreach (var part in familyData.Parts)
                            {
                                _labelStyleManager.ApplyLabelStyle(part, familyData.SelectedLabelStyle, false);
                                updatedCount++;
                            }
                        }
                        // For "Replace Existing Labels", use current style filtering
                        else
                        {
                            // If "All Styles" is selected, process all parts
                            if (familyData.CurrentSelectedStyle == "All Styles")
                            {
                                foreach (var part in familyData.Parts)
                                {
                                    _labelStyleManager.ApplyLabelStyle(part, familyData.SelectedLabelStyle, true);
                                    updatedCount++;
                                }
                            }
                            // Process only parts with the selected current style
                            else if (!string.IsNullOrEmpty(familyData.CurrentSelectedStyle) && 
                                     familyData.CurrentSelectedStyle != "Select Style..." &&
                                     familyData.CurrentSelectedStyle != "No Label")
                            {
                                foreach (var part in familyData.Parts)
                                {
                                    // Modify the label style application to be more specific
                                    if (familyData.CurrentStyles.Contains(familyData.CurrentSelectedStyle))
                                    {
                                        // Use a custom method to replace only the specific label style
                                        _labelStyleManager.ReplaceSpecificLabelStyle(part, 
                                            familyData.CurrentSelectedStyle, 
                                            familyData.SelectedLabelStyle);
                                        updatedCount++;
                                    }
                                }
                            }
                        }
                    }
                }

                if (totalSelected == 0)
                {
                    MessageBox.Show("Please select at least one part family to update.");
                    return;
                }

                MessageBox.Show($"Updated {updatedCount} parts with new label styles.");
                UpdateCurrentStyles();
            }
            catch (SystemException ex)
            {
                Logger.LogError("Error in ApplyStyle_Click", ex);
                MessageBox.Show($"Error applying styles: {ex.Message}");
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

        private void UpdateCurrentStyles()
        {
            foreach (PartFamilyData familyData in _partFamilies)
            {
                familyData.CurrentStyles.Clear();
                foreach (var part in familyData.Parts)
                {
                    string currentStyle = familyData.IsPipe 
                        ? _labelStyleManager.GetCurrentLabelStyle(part as Pipe)
                        : _labelStyleManager.GetCurrentLabelStyle(part as Structure);
                    
                    if (!string.IsNullOrEmpty(currentStyle))
                    {
                        familyData.AddCurrentStyle(currentStyle);
                    }
                }
            }
        }
    }
}