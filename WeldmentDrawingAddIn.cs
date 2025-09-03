// All Using Functions
#region using
using SolidWorks.Interop.sldworks;                 // Core SolidWorks COM interfaces (IModelDoc2, DrawingDoc, etc.)
using System.Linq;                                 // LINQ for convenient collection operations
using System;                                      // Basic .NET types and utilities
using System.Runtime.InteropServices;              // COM interop attributes (ComVisible, Guid)
using Xarial.XCad.Base.Attributes;                 // Attributes like [Title] for UI naming
using Xarial.XCad.SolidWorks;                      // Base for SolidWorks add-in (SwAddInEx)
using Xarial.XCad.SolidWorks.Documents;            // Xarial document wrappers (ISwPart, ISwBody)
using Xarial.XCad.SolidWorks.Geometry;             // Geometry wrappers
using Xarial.XCad.UI.Commands;                     // Command manager helpers
using System.Windows.Forms;                        // WinForms for dialogs and custom input form

#region using
// Aliases for SolidWorks enums and types to improve readability
using swConfigurationOptions2_e = SolidWorks.Interop.swconst.swConfigurationOptions2_e;
using swDocumentTypes_e = SolidWorks.Interop.swconst.swDocumentTypes_e;
using swDwgPaperSizes_e = SolidWorks.Interop.swconst.swDwgPaperSizes_e;
using swDwgTemplates_e = SolidWorks.Interop.swconst.swDwgTemplates_e;
using View = SolidWorks.Interop.sldworks.View;
using swExportDataFileType_e = SwConst.swExportDataFileType_e;
using swExportDataSheetsToExport_e = SwConst.swExportDataSheetsToExport_e;
using swSaveAsVersion_e = SwConst.swSaveAsVersion_e;
using swSaveAsOptions_e = SwConst.swSaveAsOptions_e;
#endregion
#endregion

namespace Solidworks_Addin
{
    // Enumeration listing all commands exposed by the add-in.
    // [Title] attributes define button captions in the SolidWorks UI.
    public enum RSRG_Automation
    {
        [Title("Generate Member and Plate Drawings")]
        Generate,

        [Title("Insert Property Annotation")]
        InsertAnnotation,

        [Title("Export Drawing to PDF")]
        ExportPDF,

        [Title("Export Plate to DXF")]
        ExportDXF
    }

    #region
    // Make the add-in visible to COM and SolidWorks with fixed GUID and ProgId.
    // Title is used for UI and add-in listings.
    [ComVisible(true)]
    [Guid("726b4113-46b3-431b-850c-b2509b38f4e8")] // Unique identifier for registration
    [ProgId("MyCompany.MyAddin")]                  // Programmatic identifier
    [Title("RSRG Solidworks Automation")]          // Displayed add-in title
    #endregion
    public class DrawingAddIn : SwAddInEx
    {
        /// <summary>
        /// Entry point when SolidWorks connects the add-in.
        /// Registers the command group and hooks button events.
        /// </summary>
        public override void OnConnect()
        {
            CommandManager
                .AddCommandGroup<RSRG_Automation>() // Create commands for enum values
                .CommandClick += OnCommandClicked;  // Subscribe to click handler
        }

        /// <summary>
        /// Dispatches command actions based on which button was clicked.
        /// </summary>
        private void OnCommandClicked(RSRG_Automation cmd)
        {
            switch (cmd)
            {
                case RSRG_Automation.Generate:
                    // Button to create Member and Plate Drawings
                    GenerateMemPlaDrawing();
                    break;

                case RSRG_Automation.InsertAnnotation:
                    // Insert Annotation into Drawing Files
                    InsertAnnotation();
                    break;

                case RSRG_Automation.ExportPDF:
                    // Export PDF
                    ExportPDF();
                    break;

                case RSRG_Automation.ExportDXF:
                    // Export DXF Files (not implemented yet)
                    ExportDXF();
                    break;
            }
        }

        /// <summary>
        /// Creates a configuration per cutlist body and generates drawing sheets
        /// with standard views referencing each configuration.
        /// </summary>
        private void GenerateMemPlaDrawing()
        {
            // Check to see if file open is an active part
            var part = Application.Documents.Active as ISwPart;
            if (part == null)
            {
                Application.ShowMessageBox("Active document is not a part.");
                return;
            }

            // Base SOLIDWORKS model pointer and path of the active part
            IModelDoc2 swModel = (IModelDoc2)part.Model;
            string partPath = swModel.GetPathName();

            // Counter and store for created configuration names
            int configIndex = 0;
            var configs = new System.Collections.Generic.List<string>();

            // Iterate each cutlist item from the active configuration
            foreach (var cutList in part.Configurations.Active.CutLists)
            {
                var firstBody = cutList.Bodies.FirstOrDefault();
                if (firstBody == null) continue; // Skip empty cutlists

                // Create a unique configuration name for this cutlist
                string configName = $"Config_CutList_{configIndex}";
                configIndex++;

                // Create configuration only if it doesn’t exist
                var existingConfigs = (string[])swModel.GetConfigurationNames();
                if (!existingConfigs.Contains(configName))
                {
                    swModel.AddConfiguration3(
                        configName,
                        $"Configuration for cut list {configIndex}",
                        "",
                        (int)swConfigurationOptions2_e.swConfigOption_DontShowPartsInBOM // Hide from BOM
                    );
                }

                // Activate target configuration
                swModel.ShowConfiguration2(configName);

                // --- Hide all bodies in the part ---
                foreach (var body in part.Bodies)
                {
                    var b = ((ISwBody)body).Body;
                    b.Select2(false, null);
                    swModel.HideSolidBody();
                    swModel.ClearSelection2(true);
                }

                // --- Show only the first body of the current cutlist ---
                var swTargetBody = ((ISwBody)firstBody).Body;
                swTargetBody.Select2(false, null);
                swModel.ShowSolidBody();
                swModel.ClearSelection2(true);

                // Rebuild to update graphics and views
                swModel.ForceRebuild3(true);

                // Track configuration for drawing generation
                configs.Add(configName);
            }

            // Error check if no valid cutlists found
            if (configs.Count == 0)
            {
                Application.ShowMessageBox("No valid cutlists found (all excluded or empty).");
                return;
            }

            // --- Drawing Creation ---
            // Template path for the A3 drawing template (adjust as needed)
            string Template = @"C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2025\templates\RSRG A3.DRWDOT";

            // Create a new drawing document using the specified template
            var swDraw = Application.Sw.INewDocument(
                Template,
                (int)swDwgPaperSizes_e.swDwgPaperA3size,
                0, 0) as DrawingDoc;

            // Track the number of drawing sheets created
            int sheetNo = 0;

            // Create one sheet per configuration and place three standard views
            foreach (string cfgName in configs)
            {
                sheetNo++;
                string sheetName = $"Sheet_{sheetNo}";

                if (sheetNo > 1)
                {
                    // Add subsequent sheets (use template border)
                    swDraw.NewSheet4(
                        sheetName,
                        (int)swDwgPaperSizes_e.swDwgPaperA3size,
                        (int)swDwgTemplates_e.swDwgTemplateCustom,
                        1.0, 10.0, false, Template,
                        0, 0, "", 0, 0, 0, 0, 0, 0
                    );
                }
                else
                {
                    // Setup first sheet (activate title block and scale)
                    swDraw.SetupSheet5(
                        sheetName,
                        (int)swDwgPaperSizes_e.swDwgPaperA3size,
                        (int)swDwgTemplates_e.swDwgTemplateCustom,
                        1.0, 10.0, true, Template,
                        0.42, 0.297, "Default", false);
                }

                // Cast to ModelDoc2 to create views
                ModelDoc2 swDrawModel = (ModelDoc2)swDraw;

                // Adds Front view (based on global coordinate system)
                View FrontView = swDraw.CreateDrawViewFromModelView3(partPath, "*Front", 0.15, 0.15, 0);
                if (FrontView != null) { FrontView.ReferencedConfiguration = cfgName; }

                // Adds Right view
                View RightView = swDraw.CreateDrawViewFromModelView3(partPath, "*Right", 0.25, 0.15, 0);
                if (RightView != null) RightView.ReferencedConfiguration = cfgName;

                // Adds Top view
                View TopView = swDraw.CreateDrawViewFromModelView3(partPath, "*Top", 0.15, 0.25, 0);
                if (TopView != null) TopView.ReferencedConfiguration = cfgName;
            }
        }

        /// <summary>
        /// Opens a form to collect annotation inputs and inserts a composed note
        /// at a fixed position on the active drawing sheet.
        /// </summary>
        private void InsertAnnotation()
        {
            // Get the active SolidWorks document and confirm it is a drawing
            IModelDoc2 swModel = (IModelDoc2)Application.Sw.ActiveDoc;
            if (swModel == null || swModel.GetType() != (int)swDocumentTypes_e.swDocDRAWING)
            {
                Application.ShowMessageBox("Please open a drawing first.");
                return;
            }

            // --- Show the input form ---
            using (UserInputForm form = new UserInputForm())
            {
                if (form.ShowDialog() != DialogResult.OK)
                {
                    Application.ShowMessageBox("Operation cancelled or no input provided.");
                    return;
                }

                // --- Build the note text using user input ---
                string GenText =
                    "NOTES: \n\n" +
                    "GENERAL\n\n" +
                    $" 1. QUANTITY SHOWN FOR {form.QuantityRequired} UNITS\n";

                // Manufacturing Method notes
                switch (form.ManMethod)
                {
                    case "None":
                        break;

                    case "Fabrication":
                        GenText += "\n FABRICATION\n\n" +
                                   " 1. DO NOT SCALE DIMENSIONS \n" +
                                   " 2. ALL DIMENSIONS ARE IN MILIMITERS \n" +
                                   " 3. ALL EDGES ARE TO BE BROKEN AND FREE OF BURS \n" +
                                   " 4. GENERAL TOLERANCE: 1mm \n" +
                                   " 5. WORKMANSHIP AND MATERIALS TO BE IAW \n" +
                                   " AS3990 & AS1554 \n";
                        break;

                    case "Lasercutting":
                        GenText += "\n LASERCUTTING \n\n" +
                                   " 1. DO NOT SCALE DIMENSIONS \n" +
                                   " 2. ALL DIMENSIONS ARE IN MILLIMETERS \n" +
                                   " 3. GENERAL TOLERANCE: 0.1mm to 0.3mm \n" +
                                   " 4. ANGULAR WORKMANSHIP AND MATERIALS TO BE IAW \n" +
                                   " AS3990 & AS1554 \n";
                        break;

                    case "Machined Component":
                        GenText += "\n FABRICATION\n\n" +
                                   " 1. DO NOT SCALE DIMENSIONS \n" +
                                   " 2. ALL DIMENSIONS ARE IN MILIMITERS \n" +
                                   " 3. ALL EDGES ARE TO BE BROKEN AND FREE OF BURS \n" +
                                   " 4. GENERAL TOLERANCE: 0.01mm \n" +
                                   " 5. WORKMANSHIP AND MATERIALS TO BE IAW \n" +
                                   " AS3990 & AS1554 \n";
                        break;
                }

                // Material/Section type notes
                switch (form.TypeUsed)
                {
                    case "None":
                        break;

                    case "Hollow Section":
                        GenText += "\n STEEL - HOLLOW SECTION \n\n" +
                                   " 1. EDGES TO BE ROUNDED TO A RADIUS OF 1mm UNO \n" +
                                   " 2. ALL STEEL SHALL BE IAW: \n" +
                                   " a. AS/NZS1163 - COLD FORMED STRUCTURAL STEEL HOLLOW SECTIONS, MIN GRADE 350 \n";
                        break;

                    case "Hot Rolled Section":
                        GenText += "\n STEEL - SECTION \n\n" +
                                   " 1. EDGES TO BE ROUNDED TO A RADIUS OF 1mm UNO \n" +
                                   " 2. ALL STEEL SHALL BE IAW: \n" +
                                   " a. AS/NZS3679.1 - HOT ROLLED BARS AND SECTION, MIN GRADE 300 \n";
                        break;

                    case "Machined Component":
                        GenText += "\n STEEL - MACHINED COMPONENT \n\n" +
                                   " 1. EDGES TO BE ROUNDED TO A RADIUS OF 1mm UNO \n" +
                                   " 2. ALL STEEL SHALL BE IAW: \n" +
                                   " a. AS 1020, MIN YIELD: 403 MPa \n";
                        break;

                    case "Plate":
                        GenText += "\n STEEL - PLATE \n\n" +
                                   " 1. EDGES TO BE ROUNDED TO A RADIUS OF 1mm UNO \n" +
                                   " 2. ALL STEEL SHALL BE IAW: \n" +
                                   " a. AS 3678 - PLATES, MIN GRADE 350 \n";
                        break;
                }

                // Welding notes if applicable
                if (form.Welding == "Yes")
                {
                    GenText += "\n WELDING \n\n" +
                               " 1. ALL WELDS SHALL CONFORM TO AS1554.1 SP \n" +
                               " 2. 100% VISUAL INSPECTION, 10% MPI \n" +
                               " 3. NOMINAL TENSILE STRENGTH OF WELDS GREATER THAN \n" +
                               " THAN PARENT MATERIAL \n" +
                               $" 4. ALL WELDS SHALL BE {form.Weld}mm CONTINOUS FILLETS UNO \n" +
                               $" 5. {form.Weld}mm CONTINOUS FILLET WELDS TO REFERENCE \n" +
                               " WELD PROCEDURE RE 001 \n" +
                               " 6. ALL BUTT WELDS SHALL BE FULL PENERATION UNO \n" +
                               " 7. BUTT WELDS TO REFERENCE RE 003 \n" +
                               " 8. ALL WELDS ARE TO BE SHOP WELDS UNO \n" +
                               " 9. ELECTRODES TO BE E49XX ELECTRODES \n";
                }

                // Surface treatment notes
                switch (form.SurTreat)
                {
                    case "None":
                        break;

                    case "Galvanised":
                        GenText += "\n SURFACE TREATMENT \n\n" +
                                   " 1. ALL STEELWORK SHALL BE HOT DIP GALVANISED TO AS4680";
                        break;

                    case "ColdGal":
                        GenText += "\n SURFACE TREATMENT \n\n" +
                                   " 1. ALL STEELWORK TO BE SPRAYED WITH COLD GAL";
                        break;

                    case "Painted":
                        GenText += "\n SURFACE TREATMENT \n\n" +
                                   " 1. ALL STEELWORK TO BE PAINTED WITH <SPECIFY COLOUR> \n";
                        break;

                    case "Metal Spray":
                        GenText += "\n SURFACE TREATMENT \n\n" +
                                   " 1. ALL STEELWORK TO BE SPARYED WITH METAL SPRAYED \n";
                        break;
                }

                // --- Insert note and position it on the sheet ---
                Note swNote = (Note)swModel.InsertNote(GenText);
                if (swNote == null)
                {
                    Application.ShowMessageBox("Failed to insert note.");
                    return;
                }

                // Use IAnnotation to set 2D position in meters
                Annotation swAnn = swNote.GetAnnotation() as Annotation;
                if (swAnn != null)
                {
                    swAnn.SetPosition(0.2, 0.2, 0); // X=0.2m, Y=0.2m on the sheet
                }

                // Clear selection to tidy up state
                swModel.ClearSelection2(true);
            }
        }

        /// <summary>
        /// Exports the active drawing to a multi-sheet PDF using ExportPdfData.
        /// Displays a save dialog and reports status to the user.
        /// </summary>
        private void ExportPDF()
        {
            // Get active document and ensure it is a drawing
            IModelDoc2 swModel = (IModelDoc2)Application.Sw.ActiveDoc;
            if (swModel == null || swModel.GetType() != (int)swDocumentTypes_e.swDocDRAWING)
            {
                Application.ShowMessageBox("Please open a drawing before exporting.");
                return;
            }

            // --- Show Save File Dialog to choose PDF file path ---
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Title = "Save PDF As";
                saveFileDialog.Filter = "PDF Files (*.pdf)|*.pdf";
                saveFileDialog.DefaultExt = "pdf";
                // Suggest filename based on drawing title
                saveFileDialog.FileName = System.IO.Path.GetFileNameWithoutExtension(swModel.GetTitle()) + ".pdf";

                if (saveFileDialog.ShowDialog() != DialogResult.OK)
                {
                    Application.ShowMessageBox("Export cancelled.");
                    return;
                }

                // Target PDF path selected by user
                string pdfPath = saveFileDialog.FileName;

                // --- Prepare export data for PDF ---
                ModelDocExtension swExt = swModel.Extension;
                int errors = 0;
                int warnings = 0;

                // Create ExportPdfData object via SOLIDWORKS API
                ExportPdfData pdfData = (ExportPdfData)Application.Sw.GetExportFileData(
                    (int)swExportDataFileType_e.swExportPdfData);

                // Configure options: export all sheets, no 3D, do not auto-open
                pdfData.ViewPdfAfterSaving = false;
                pdfData.ExportAs3D = false;
                pdfData.SetSheets(
                    (int)swExportDataSheetsToExport_e.swExportData_ExportAllSheets, null);

                // --- Save as PDF silently with provided options ---
                bool status = swExt.SaveAs(
                    pdfPath,
                    (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                    (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                    pdfData,
                    ref errors,
                    ref warnings
                );

                // Report result to user
                if (status)
                {
                    Application.ShowMessageBox($"All sheets exported to PDF:\n{pdfPath}");
                }
                else
                {
                    Application.ShowMessageBox($"Failed to export PDF. Errors: {errors}, Warnings: {warnings}");
                }
            }
        }

        /// <summary>
        /// Placeholder for DXF export logic.
        /// Implement plate-specific DXF creation here.
        /// </summary>
        private void ExportDXF()
        {
        }

        /// <summary>
        /// Modal form to collect inputs for generating the general notes block.
        /// Exposes properties for retrieved values after OK is pressed.
        /// </summary>
        public class UserInputForm : Form
        {
            // Public read-only properties expose inputs to the caller
            public string QuantityRequired { get; private set; }
            public string Weld { get; private set; }
            public string ManMethod { get; private set; }
            public string TypeUsed { get; private set; }
            public string Welding { get; private set; }
            public string SurTreat { get; private set; }

            // UI controls
            private TextBox txtQuantityRequired;
            private TextBox txtWeld;
            private ComboBox cmbManMethod;
            private ComboBox cmbTypeUsed;
            private ComboBox cmbWelding;
            private ComboBox cmbSurTreat;
            private Button btnOk;

            public UserInputForm()
            {
                // Initialize form parameters
                this.Text = "Enter Annotation Details";
                this.Width = 300;
                this.Height = 460;

                // Layout manager (simple two-column layout)
                var layout = new TableLayoutPanel();
                layout.Dock = DockStyle.Fill;
                layout.RowCount = 8;
                layout.ColumnCount = 2;
                layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
                layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70));

                // Text Box for Quantity Required
                var lblQuantity = new Label() { Text = "Quantity Required", Left = 10, Top = 10, Width = 260 };
                txtQuantityRequired = new TextBox { Left = 10, Top = 30, Width = 260 };

                // Dropdown for Manufacturing Method
                var lblManMethod = new Label() { Text = "select the type of manufacturing Method Used", Left = 10, Top = 80, Width = 260 };
                cmbManMethod = new ComboBox { Left = 10, Top = 100, Width = 260 };
                cmbManMethod.Items.AddRange(new string[] { "None", "Fabrication", "Lasercutting", "Machined Component" });
                cmbManMethod.DropDownStyle = ComboBoxStyle.DropDownList;
                cmbManMethod.SelectedIndex = 0; // default

                // Dropdown for Section Type / Material
                var lblMaterialUsd = new Label() { Text = "select the type of Section Used", Left = 10, Top = 150, Width = 260 };
                cmbTypeUsed = new ComboBox { Left = 10, Top = 170, Width = 260 };
                cmbTypeUsed.Items.AddRange(new string[] { "None", "Hollow Section", "Hot Rolled Section", "Machined Component", "Plate" });
                cmbTypeUsed.DropDownStyle = ComboBoxStyle.DropDownList;
                cmbTypeUsed.SelectedIndex = 0; // default

                // Weld thickness
                var lblWeld = new Label() { Text = "What is the thickness of the weld", Left = 10, Top = 220, Width = 260 };
                txtWeld = new TextBox { Left = 10, Top = 250, Width = 260 };

                // Is Welding used
                var lblWelding = new Label() { Text = "Is this for a weldment", Left = 10, Top = 300, Width = 260 };
                cmbWelding = new ComboBox { Left = 10, Top = 330, Width = 260 };
                cmbWelding.Items.AddRange(new string[] { "Yes", "No" });
                cmbWelding.DropDownStyle = ComboBoxStyle.DropDownList;
                cmbWelding.SelectedIndex = 0; // default

                // Surface Treatment selection
                var lblSurTreat = new Label() { Text = "What Surface Treatment is used", Left = 10, Top = 380, Width = 260 };
                cmbSurTreat = new ComboBox { Left = 10, Top = 400, Width = 260 };
                cmbSurTreat.Items.AddRange(new string[] { "None", "Galvanised", "ColdGal", "Painted", "Metal Spray" });
                cmbSurTreat.DropDownStyle = ComboBoxStyle.DropDownList;
                cmbSurTreat.SelectedIndex = 0; // default

                // OK button confirms and transfers values to properties
                btnOk = new Button { Text = "OK", Left = 10, Top = 440, Width = 100 };
                btnOk.Click += (sender, e) =>
                {
                    // Persist user selections to properties
                    QuantityRequired = txtQuantityRequired.Text;
                    ManMethod = cmbManMethod.SelectedItem.ToString();
                    TypeUsed = cmbTypeUsed.SelectedItem.ToString();
                    Weld = txtWeld.Text;
                    Welding = cmbWelding.SelectedItem.ToString();
                    SurTreat = cmbSurTreat.SelectedItem.ToString();

                    // Close with OK result to signal success
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                };

                // Add controls to the form
                Controls.Add(lblQuantity);
                Controls.Add(txtQuantityRequired);
                Controls.Add(lblManMethod);
                Controls.Add(cmbManMethod);
                Controls.Add(lblMaterialUsd);
                Controls.Add(cmbTypeUsed);
                Controls.Add(lblWeld);
                Controls.Add(txtWeld);
                Controls.Add(lblWelding);
                Controls.Add(cmbWelding);
                Controls.Add(lblSurTreat);
                Controls.Add(cmbSurTreat);
                Controls.Add(btnOk);
            }
        }
    }
}