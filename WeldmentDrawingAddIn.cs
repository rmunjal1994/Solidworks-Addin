// All Using Functions 
#region
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using System.Linq;
using System;
using Xarial.XCad;
using System.Drawing;
using System.Runtime.InteropServices;
using Xarial.XCad.Base.Attributes;
using Xarial.XCad.SolidWorks;
using Xarial.XCad.SolidWorks.Documents;
using Xarial.XCad.SolidWorks.Features;
using Xarial.XCad.SolidWorks.Geometry;
using Xarial.XCad.UI.Commands;
using System.Windows.Forms;
using Microsoft.VisualBasic;
#region
using swConfigurationOptions2_e = SolidWorks.Interop.swconst.swConfigurationOptions2_e;
using swDocumentTypes_e = SolidWorks.Interop.swconst.swDocumentTypes_e;
using swDwgPaperSizes_e = SolidWorks.Interop.swconst.swDwgPaperSizes_e;
using swDwgTemplates_e = SolidWorks.Interop.swconst.swDwgTemplates_e;
using View = SolidWorks.Interop.sldworks.View;
using SwConst;
using Xarial.XCad.Documents;
using swExportDataFileType_e = SwConst.swExportDataFileType_e;
using swExportDataSheetsToExport_e = SwConst.swExportDataSheetsToExport_e;
using swSaveAsVersion_e = SwConst.swSaveAsVersion_e;
using swSaveAsOptions_e = SwConst.swSaveAsOptions_e;
using System.Diagnostics.Tracing;
#endregion
#endregion

namespace Solidworks_Addin
{
    public enum Commands_e                                                      //Initilise Buttons within solidworks 
    {
        [Title("Generate Member and Plate Drawings")] Generate,     

        [Title("Insert Property Annotation")] InsertAnnotation,

        [Title("Export Drawing to PDF")] ExportPDF,

        [Title("Export Plate to DXF")] ExportDXF
    }

    #region
    [ComVisible(true)]
    [Guid("726b4113-46b3-431b-850c-b2509b38f4e8")]
    [ProgId("MyCompany.MyAddin")]
    [Title("RSRG Solidworks Automation")]
    #endregion

    public class DrawingAddIn : SwAddInEx
    {
       

        public override void OnConnect()                            //Initiliase Addin on connect 
        {
            CommandManager
                .AddCommandGroup<Commands_e>()
                .CommandClick += OnCommandClicked;
        }

        private void OnCommandClicked(Commands_e cmd)
        {
            switch (cmd)
            {
                case Commands_e.Generate:
                    GenerateMemPlaDrawing();
                    break;

                case Commands_e.InsertAnnotation:
                    InsertAnnotation();
                    break;

                case Commands_e.ExportPDF:
                    ExportPDF(); // hardcoded file path for save locataion for PDF File 
                    break;

                case Commands_e.ExportDXF:
                    ExportDXF();
                    break;

            }
        }
       
        private void GenerateMemPlaDrawing()
        {
            var part = Application.Documents.Active as ISwPart;     // check to see if file open is an active part 
            if (part == null)
            {
                Application.ShowMessageBox("Active document is not a part.");
                return;
            }
            IModelDoc2 swModel = (IModelDoc2)part.Model;        
            string partPath = swModel.GetPathName();

            int configIndex = 0;                                                    // counter for sheet name number to be used to track number of configs created 
            var configs = new System.Collections.Generic.List<string>();            //list to save all required configs 

            // Iterate each cutlist item
            foreach (var cutList in part.Configurations.Active.CutLists)            
            {

                var firstBody = cutList.Bodies.FirstOrDefault();
                if (firstBody == null) continue;

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
                        (int)swConfigurationOptions2_e.swConfigOption_DontShowPartsInBOM
                    );
                }

                swModel.ShowConfiguration2(configName);

                // --- Hide all bodies ---
                foreach (var body in part.Bodies)
                {
                    var b = ((ISwBody)body).Body;
                    b.Select2(false, null);
                    swModel.HideSolidBody();
                    swModel.ClearSelection2(true);
                }

                // --- Show only first body of cutlist ---
                var swTargetBody = ((ISwBody)firstBody).Body;
                swTargetBody.Select2(false, null);
                swModel.ShowSolidBody();
                swModel.ClearSelection2(true);

                swModel.ForceRebuild3(true);
                configs.Add(configName);
            }

            if (configs.Count == 0)                                                                    // error check to determine if cutlist selected 
            {
                Application.ShowMessageBox("No valid cutlists found (all excluded or empty).");
                return;
            }

            // --- Drawing Creation ---
            string Template = @"C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2025\templates\RSRG A3.DRWDOT"; // template location for drawing sheet template 
            var swDraw = Application.Sw.INewDocument(Template,
                         (int)swDwgPaperSizes_e.swDwgPaperA3size, 0, 0) as DrawingDoc;

            int sheetNo = 0;                                                                        // track number of drawing sheet created 

            foreach (string cfgName in configs)
            {
                sheetNo++;
                string sheetName = $"Sheet_{sheetNo}";

                if (sheetNo > 1)
                {
                    swDraw.NewSheet4(sheetName,                                                    //setup new sheet 
                        (int)swDwgPaperSizes_e.swDwgPaperA3size,
                        (int)swDwgTemplates_e.swDwgTemplateCustom,
                        1.0, 10.0, false, Template, 0, 0, "", 0, 0, 0, 0, 0, 0
                    );
                }
                else
                {
                    swDraw.SetupSheet5(sheetName,                                                  //Setup First Sheet 
                        (int)swDwgPaperSizes_e.swDwgPaperA3size,
                        (int)swDwgTemplates_e.swDwgTemplateCustom,
                        1.0, 10.0, true, Template, 0.42, 0.297, "Default", false);
                }

                ModelDoc2 swDrawModel = (ModelDoc2)swDraw;


                View FrontView = swDraw.CreateDrawViewFromModelView3(partPath, "*Front", 0.15, 0.15, 0); // Adds front View(based of global coordinate system) 
                if (FrontView != null)
                {
                    FrontView.ReferencedConfiguration = cfgName;
                }

                View RightView = swDraw.CreateDrawViewFromModelView3(partPath, "*Right", 0.25, 0.15, 0);// Adds Right View(based of global coordinate system)
                if (RightView != null) RightView.ReferencedConfiguration = cfgName;

                View TopView = swDraw.CreateDrawViewFromModelView3(partPath, "*Top", 0.15, 0.25, 0);// Adds Top View(based of global coordinate system)
                if (TopView != null) TopView.ReferencedConfiguration = cfgName;
            }
        }

        private void InsertAnnotation()
        {
            // Get the active SolidWorks document
            IModelDoc2 swModel = (IModelDoc2)Application.Sw.ActiveDoc;
            if (swModel == null || swModel.GetType() != (int)swDocumentTypes_e.swDocDRAWING)  // checks if open file is drawing 
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
                    "NOTES: \n" + "\n" +
                    "GENERAL\n" + "\n" + 
                    $"   1. QUANTITY SHOWN FOR {form.QuantityRequired} UNITS\n";

                switch (form.ManMethod)                 //--For Manufacturing Method--
                {
                    case "None":
                        break;
                    case "Fabrication":                                                 // reference dropdown option 
                        GenText += "\n FABRICATION\n" + "\n" +                          // "\n" to create new line 
                                    "   1. DO NOT SCALE DIMENSIONS \n" + 
                                    "   2. ALL DIMENSIONS ARE IN MILIMITERS \n" + 
                                    "   3. ALL EDGES ARE TO BE BROKEN AND FREE OF BURS \n" +
                                    "   4. GENERAL TOLERANCE: 1mm \n"+
                                    "   5. WORKMANSHIP AND MATERIALS TO BE IAW \n" + 
                                    "       AS3990 & AS1554 \n";
                        break;
                    
                    case "Lasercutting":
                        GenText += "\n LASERCUTTING \n" + "\n" +
                                   "    1. DO NOT SCALE DIMENSIONS \n" + 
                                   "    2. ALL DIMENSIONS ARE IN MILLIMETERS \n" +
                                   "    3. GENERAL TOLERANCE: 0.1mm to 0.3mm \n" +
                                   "    4. ANGULAR WORKMANSHIP AND MATERIALS TO BE IAW \n" + 
                                   "            AS3990 & AS1554 \n";
                        break;

                    case "Machined Component":
                        GenText += "\n FABRICATION\n" + "\n" +
                                    "   1. DO NOT SCALE DIMENSIONS \n" +
                                    "   2. ALL DIMENSIONS ARE IN MILIMITERS \n" +
                                    "   3. ALL EDGES ARE TO BE BROKEN AND FREE OF BURS \n" +
                                    "   4. GENERAL TOLERANCE: 0.01mm \n" +
                                    "   5. WORKMANSHIP AND MATERIALS TO BE IAW \n" + 
                                    "           AS3990 & AS1554 \n";
                        break;
                };

                switch (form.TypeUsed)
                {
                    case "None":
                        break;
                    case "Hollow Section":
                        GenText += "\n STEEL - HOLLOW SECTION \n" + "\n" +
                                   "    1. EDGES TO BE ROUNDED TO A RADIUS OF 1mm UNO \n" + 
                                   "    2. ALL STEEL SHALL BE IAW: \n" + 
                                   "        a. AS/NZS1163 - COLD FORMED STRUCTURAL STEEL HOLLOW SETCIONS, MIN GRADE 350 \n";
                            break;
                    case "Hot Rolled Section":
                        GenText += "\n STEEL - SECTION \n" + "\n" +
                                   "    1. EDGES TO BE ROUNDED TO A RADIUS OF 1mm UNO \n" +
                                   "    2. ALL STEEL SHALL BE IAW: \n" +
                                   "        a. AS/NZS3679.1 - HOT ROLLED BARS AND SECTION, MIN GRADE 300 \n";
                            break;
                    case "Machined Component":
                        GenText += "\n STEEL - MACHINED COMPONENT \n" + "\n" +
                                   "    1. EDGES TO BE ROUNDED TO A RADIUS OF 1mm UNO \n" +
                                   "    2. ALL STEEL SHALL BE IAW: \n" +
                                   "        a. AS 1020, MIN YIELD: 403 MPa \n";
                        break;
                    case "Plate":
                        GenText += "\n STEEL - PLATE \n" + "\n" +
                                   "    1. EDGES TO BE ROUNDED TO A RADIUS OF 1mm UNO \n" +
                                   "    2. ALL STEEL SHALL BE IAW: \n" +
                                   "        a. AS 3678 - PLATES, MIN GRADE 350 \n";
                        break;
                };
                if (form.Welding == "Yes")
                {
                    GenText += "\n WELDING \n" + "\n" +
                               "    1. ALL WELDS SHALL CONFORM TO AS1554.1 SP \n" +
                               "    2. 100% VISUAL INSPECTION, 10% MPI \n" +
                               "    3. NOMINAL TENSILE STRENGTH OF WELDS GREATER THAN \n" +
                               "       THAN PARENT MATERIAL \n" +
                               $"   4. ALL WELDS SHALL BE {form.Weld}mm CONTINOUS FILLETS UNO \n" +
                               $"   5. {form.Weld}mm CONTINOUS FILLET WELDS TO REFERENCE \n" +
                               "        WELD PROCEDURE RE 001 \n" +
                               "    6. ALL BUTT WELDS SHALL BE FULL PENERATION UNO \n" +
                               "    7. BUTT WELDS TO REFERENCE RE 003 \n" +
                               "    8. ALL WELDS ARE TO BE SHOP WELDS UNO \n" +
                               "    9. ELECTRODES TO BE E49XX ELECTRODES \n";
                }
                else
                {
                    GenText += "";
                }

            switch (form.SurTreat)
                {
                    case "None":
                        break;

                    case "Galvanised":

                        GenText += "\n SURFACE TREATMENT \n" + "\n" +
                                   "    1. ALL STEELWORK SHALL BE HOT DIP GALVANISED TO AS4680";
                        break;

                    case "ColdGal":
                        GenText += "\n SURFACE TREATMENT \n" + "\n" +
                                   "    1. ALL STEELWORK TO BE SPRAYED WITH COLDGAL";
                        break;

                    case "Painted":
                        GenText += "\n SURFACE TREATMENT \n" + "\n" +
                                   "    1. ALL STEELWORK TO BE PAINTED WITH <SPECIFY COLOUR> \n";
                        break;

                    case "Metal Spray":
                        GenText += "\n SURFACE TREATMENT \n" + "\n" +
                                   "    1. ALL STEELWORK TO BE SPARYED WITH METAL SPRAYED \n";
                        break;

                }

                    // --- Insert note at sheet coordinates (0.2, 0.2 meters) ---
                    Note swNote = (Note)swModel.InsertNote(GenText);
                if (swNote == null)
                {
                    Application.ShowMessageBox("Failed to insert note.");
                    return;
                }

                // Use IAnnotation to set position
                Annotation swAnn = swNote.GetAnnotation() as Annotation;
                if (swAnn != null)
                {
                    swAnn.SetPosition(0.2, 0.2, 0); // SolidWorks native API method
                }

                swModel.ClearSelection2(true);
            }
        }

        private void ExportPDF()
        {

            // Get active document
            IModelDoc2 swModel = (IModelDoc2)Application.Sw.ActiveDoc;

            if (swModel == null || swModel.GetType() != (int)swDocumentTypes_e.swDocDRAWING)
            {
                Application.ShowMessageBox("Please open a drawing before exporting.");
                return;
            }

            // --- Show Save File Dialog ---
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Title = "Save PDF As";
                saveFileDialog.Filter = "PDF Files (*.pdf)|*.pdf";
                saveFileDialog.DefaultExt = "pdf";
                saveFileDialog.FileName = System.IO.Path.GetFileNameWithoutExtension(swModel.GetTitle()) + ".pdf"; // default name

                if (saveFileDialog.ShowDialog() != DialogResult.OK)
                {
                    Application.ShowMessageBox("Export cancelled.");
                    return;
                }

                string pdfPath = saveFileDialog.FileName;

                // --- Get extension object ---
                ModelDocExtension swExt = swModel.Extension;
                int errors = 0;
                int warnings = 0;

                // Create ExportPdfData object
                ExportPdfData pdfData = (ExportPdfData)Application.Sw.GetExportFileData((int)swExportDataFileType_e.swExportPdfData);

                // Export ALL sheets
                pdfData.ViewPdfAfterSaving = false;
                pdfData.ExportAs3D = false;
                pdfData.SetSheets((int)swExportDataSheetsToExport_e.swExportData_ExportAllSheets, null);

                // Save as PDF
                bool status = swExt.SaveAs(
                    pdfPath,
                    (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                    (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                    pdfData,
                    ref errors,
                    ref warnings
                );

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

        private void ExportDXF()
        {

        }

        public class UserInputForm : Form                               // method to create a form for note generation 
        {

            public string QuantityRequired { get; private set; }
            public string Weld { get; private set; }
            public string ManMethod { get; private set; }
            public string TypeUsed { get; private set; }
            public string Welding { get; private set; }
            public string SurTreat { get; private set; }


            private TextBox txtQuantityRequired;
            private TextBox txtWeld;
            private ComboBox cmbManMethod;
            private ComboBox cmbTypeUsed;
            private ComboBox cmbWelding;
            private ComboBox cmbSurTreat;
            private Button btnOk;

            public UserInputForm()
            {
                // Initilise Form Parameters 
                this.Text = "Enter Annotation Details";
                this.Width = 300;
                this.Height = 500;

                var layout = new TableLayoutPanel();
                layout.Dock = DockStyle.Fill;
                layout.RowCount = 8;   
                layout.ColumnCount = 2;
                layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
                layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70));

                //Text Box for Quantity Required 
                var lblQuantity = new Label() { Text = "Quantity Required", Left = 10, Top = 10, Width = 260 };     // Heading for Title 
                txtQuantityRequired = new TextBox { Left = 10, Top = 30, Width = 260 };                             //{Location within form}

                // --- Dropdown for Type of Manufacturing Process Used ---
                var lblManMethod = new Label() { Text = "select the type of manufacturing Method Used", Left = 10, Top = 50, Width = 260 };
                cmbManMethod = new ComboBox { Left = 10, Top = 70, Width = 260 };
                cmbManMethod.Items.AddRange(new string[] { "None","Fabrication", "Lasercutting", "Machined Component" }); // List within dropdown 
                cmbManMethod.DropDownStyle = ComboBoxStyle.DropDownList;
                cmbManMethod.SelectedIndex = 0; // default selection

                // --- Dropdown for Type of Material Used  ---
                var lblMaterialUsd = new Label() { Text = "select the type of Section Used", Left = 10, Top = 100, Width = 260 };
                cmbTypeUsed = new ComboBox { Left = 10, Top = 120, Width = 260 };
                cmbTypeUsed.Items.AddRange(new string[] { "None","Hollow Section", "Hot Rolled Section", "Machined Component", "Plate" });
                cmbTypeUsed.DropDownStyle = ComboBoxStyle.DropDownList;
                cmbTypeUsed.SelectedIndex = 0; // default selection


                //--Weld Thickness--
                var lblWeld = new Label() { Text = "What is the thickness of the weld", Left = 10, Top = 150, Width = 260 };
                txtWeld = new TextBox { Left = 10, Top = 170, Width = 260 };

                // --- Is Welding used  ---
                var lblWelding = new Label() { Text = "Is this for a weldment", Left = 10, Top = 200, Width = 260 };
                cmbWelding = new ComboBox { Left = 10, Top = 220 , Width = 260 };
                cmbWelding.Items.AddRange(new string[] { "Yes", "No" });
                cmbWelding.DropDownStyle = ComboBoxStyle.DropDownList;
                cmbWelding.SelectedIndex = 0; // default selection

                // --- What Surface Treatment is used  ---
                var lblSurTreat = new Label() { Text = "What Surface Treatment is used", Left = 10, Top = 240, Width = 260 };
                cmbSurTreat = new ComboBox { Left = 10, Top = 260, Width = 260 };
                cmbSurTreat.Items.AddRange(new string[] { "None", "Galvanised", "ColdGal", "Painted", "Metal Spray" });
                cmbSurTreat.DropDownStyle = ComboBoxStyle.DropDownList;
                cmbSurTreat.SelectedIndex = 0; // default selection



                btnOk = new Button { Text = "OK", Left = 10, Top = 280, Width = 100 }; 
                btnOk.Click += (sender, e) =>               // comand for implement when Ok Button is pressed within 
                {
                    QuantityRequired = txtQuantityRequired.Text;                                 
                    ManMethod = cmbManMethod.SelectedItem.ToString();
                    TypeUsed = cmbTypeUsed.SelectedItem.ToString();
                    Weld = txtWeld.Text;
                    Welding = cmbWelding.SelectedItem.ToString();
                    SurTreat = cmbSurTreat.SelectedItem.ToString();
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                };


                Controls.Add(lblQuantity);                          //Add functions within Form 
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

    



        
