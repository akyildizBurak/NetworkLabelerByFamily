using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.Civil.DatabaseServices.Styles;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.Civil.Settings;
using System.Collections;
using System.Reflection;
using Autodesk.AutoCAD.Geometry;

namespace NetworkLabeler
{
    public class LabelStyleManager
    {
        private readonly CivilDocument _civilDoc;
        private Dictionary<string, string> _labelStyleSelections;

        public LabelStyleManager(CivilDocument civilDoc)
        {
            _civilDoc = civilDoc;
            _labelStyleSelections = new Dictionary<string, string>();
        }

        public List<string> GetPipeLabelStyles()
        {
            var styleNames = new List<string>();
            using (var tr = Application.DocumentManager.MdiActiveDocument.Database.TransactionManager.StartTransaction())
            {
                var pipeLabelStyles = _civilDoc.Styles.LabelStyles.PipeLabelStyles;
                var styleList = new ArrayList();
                ListRoot(pipeLabelStyles, styleList);
                foreach (StyleInfo styleInfo in styleList)
                {
                    styleNames.Add(styleInfo.name);
                }
                tr.Commit();
            }
            return styleNames;
        }

        public List<string> GetStructureLabelStyles()
        {
            var styleNames = new List<string>();
            using (var tr = Application.DocumentManager.MdiActiveDocument.Database.TransactionManager.StartTransaction())
            {
                var structureLabelStyles = _civilDoc.Styles.LabelStyles.StructureLabelStyles;
                var styleList = new ArrayList();
                ListRoot(structureLabelStyles, styleList);
                foreach (StyleInfo styleInfo in styleList)
                {
                    styleNames.Add(styleInfo.name);
                }
                tr.Commit();
            }
            return styleNames;
        }

        private void ListRoot(object root, ArrayList styleList)
        {
            if (root == null) return;

            Type objectType = root.GetType();
            PropertyInfo[] properties = objectType.GetProperties(BindingFlags.Public | BindingFlags.DeclaredOnly | BindingFlags.Instance);

            foreach (PropertyInfo pf in properties)
            {
                if (pf.PropertyType.ToString().Contains("Collection"))
                {
                    ListCollection(objectType, pf, root, styleList);
                }
                else if (pf.PropertyType.ToString().Contains("Root"))
                {
                    object root2 = objectType.InvokeMember(pf.Name, BindingFlags.GetProperty, null, root, new object[0]);
                    if (root2 != null)
                    {
                        ListRoot(root2, styleList);
                    }
                }
            }
        }

        private void ListCollection(Type objectType, PropertyInfo pf, object myStylesRoot, ArrayList styleList)
        {
            object res = objectType.InvokeMember(pf.Name, BindingFlags.GetProperty, null, myStylesRoot, new object[0]);
            if (res == null) return;

            StyleCollectionBase scBase = res as StyleCollectionBase;
            if (scBase == null) return;

            foreach (ObjectId sbid in scBase)
            {
                StyleBase stylebase = Application.DocumentManager.MdiActiveDocument.Database.TransactionManager.GetObject(sbid, OpenMode.ForRead) as StyleBase;
                if (stylebase != null)
                {
                    StyleInfo styleinfo = new StyleInfo();
                    styleinfo.name = stylebase.Name;
                    styleinfo.type = stylebase.GetType().ToString();
                    styleinfo.parent = pf.Name;
                    styleList.Add(styleinfo);
                }
            }
        }

        public void SaveSelections(string drawingPath)
        {
            string configPath = Path.ChangeExtension(drawingPath, ".labelconfig");
            using (StreamWriter writer = new StreamWriter(configPath))
            {
                foreach (var selection in _labelStyleSelections)
                {
                    writer.WriteLine($"{selection.Key}={selection.Value}");
                }
            }
        }

        public void LoadSelections(string drawingPath)
        {
            string configPath = Path.ChangeExtension(drawingPath, ".labelconfig");
            if (!File.Exists(configPath)) return;

            _labelStyleSelections.Clear();
            using (StreamReader reader = new StreamReader(configPath))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    var parts = line.Split('=');
                    if (parts.Length == 2)
                    {
                        _labelStyleSelections[parts[0]] = parts[1];
                    }
                }
            }
        }

        public string GetSelectedLabelStyle(string familyName)
        {
            return _labelStyleSelections.ContainsKey(familyName) ? _labelStyleSelections[familyName] : null;
        }

        public void SetSelectedLabelStyle(string familyName, string labelStyleName)
        {
            _labelStyleSelections[familyName] = labelStyleName;
        }

        public void ApplyLabelStyle(Part part, string labelStyleName)
        {
            using (var tr = Application.DocumentManager.MdiActiveDocument.Database.TransactionManager.StartTransaction())
            {
                try
                {
                    if (part is Pipe pipe)
                    {
                        var pipeLabelStyles = _civilDoc.Styles.LabelStyles.PipeLabelStyles;
                        var styleList = new ArrayList();
                        ListRoot(pipeLabelStyles, styleList);
                        foreach (StyleInfo styleInfo in styleList)
                        {
                            if (styleInfo.name == labelStyleName)
                            {
                                var styleId = GetStyleIdByName(pipeLabelStyles, labelStyleName);
                                if (styleId != ObjectId.Null)
                                {
                                    PipeLabel.Create(pipe.ObjectId, 0.5, styleId);
                                }
                                break;
                            }
                        }
                    }
                    else if (part is Structure structure)
                    {
                        var structureLabelStyles = _civilDoc.Styles.LabelStyles.StructureLabelStyles;
                        var styleList = new ArrayList();
                        ListRoot(structureLabelStyles, styleList);
                        foreach (StyleInfo styleInfo in styleList)
                        {
                            if (styleInfo.name == labelStyleName)
                            {
                                var styleId = GetStyleIdByName(structureLabelStyles, labelStyleName);
                                if (styleId != ObjectId.Null)
                                {
                                    // Get the structure's location for the label
                                    var structureLocation = structure.Location;
                                    StructureLabel.Create(structure.ObjectId, styleId, structureLocation);
                                }
                                break;
                            }
                        }
                    }
                    tr.Commit();
                }
                catch (Exception ex)
                {
                    tr.Abort();
                    throw;
                }
            }
        }

        private ObjectId GetStyleIdByName(object stylesRoot, string styleName)
        {
            var styleList = new ArrayList();
            ListRoot(stylesRoot, styleList);
            foreach (StyleInfo styleInfo in styleList)
            {
                if (styleInfo.name == styleName)
                {
                    // Get the style ID from the collection
                    Type objectType = stylesRoot.GetType();
                    PropertyInfo[] properties = objectType.GetProperties(BindingFlags.Public | BindingFlags.DeclaredOnly | BindingFlags.Instance);
                    foreach (PropertyInfo pf in properties)
                    {
                        if (pf.PropertyType.ToString().Contains("Collection"))
                        {
                            object res = objectType.InvokeMember(pf.Name, BindingFlags.GetProperty, null, stylesRoot, new object[0]);
                            if (res is StyleCollectionBase scBase)
                            {
                                foreach (ObjectId sbid in scBase)
                                {
                                    StyleBase stylebase = Application.DocumentManager.MdiActiveDocument.Database.TransactionManager.GetObject(sbid, OpenMode.ForRead) as StyleBase;
                                    if (stylebase != null && stylebase.Name == styleName)
                                    {
                                        return sbid;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return ObjectId.Null;
        }
    }

    public class StyleInfo
    {
        public string parent;
        public string name;
        public string type;
        public Dictionary<string, string> paramValues;

        public StyleInfo()
        {
            paramValues = new Dictionary<string, string>();
        }
    }
} 