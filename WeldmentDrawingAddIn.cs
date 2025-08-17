// All Using FUnctions 
#region
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using System;
using System.CodeDom;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Media;
using Xarial.XCad;
using Xarial.XCad.Base;
using Xarial.XCad.Base.Attributes;
using Xarial.XCad.Documents;
using Xarial.XCad.Documents.Enums;
using Xarial.XCad.Documents.Extensions;
using Xarial.XCad.Features;
using Xarial.XCad.Geometry;
using Xarial.XCad.SolidWorks;
using Xarial.XCad.SolidWorks.Documents;
using Xarial.XCad.SolidWorks.Features;
using Xarial.XCad.SolidWorks.Geometry;
using Xarial.XCad.SolidWorks.UI;
using Xarial.XCad.UI.Commands;
using Xarial.XCad.Utils.Reflection;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;
#endregion
namespace Solidworks_Addin
{
    public enum WeldmentCommands_e
    {
        [Title("Generate Weldment Drawings")] Generate
    }

    [ComVisible(true)]
    [Title("Weldment Drawing Generator")]
    public class DrawingAddIn : SwAddInEx
    {
        public override void OnConnect()
        {
            CommandManager
                .AddCommandGroup<WeldmentCommands_e>()
                .CommandClick += OnCommandClicked;
        }

        private void OnCommandClicked(WeldmentCommands_e cmd)
        {
            var part = Application.Documents.Active as ISwPart;
            var swPart = part.Model as IPartDoc;
            IModelDoc2 swModel = (IModelDoc2)part.Model;



            if (part == null)
            {
                Application.ShowMessageBox("Active document is not a part.");
                return;
            }
            // Organise Model 
            #region
            int SheetIndex = 0;        // Collect the first body from each cutlist item 
            int bodyIndex = 1;         // Inatilise Count to catch number of bodies 

            foreach (var cutList in part.Configurations.Active.CutLists)
            {
                var firstBody = cutList.Bodies.FirstOrDefault();
                

                var swBody = ((ISwBody)firstBody).Body;
                var swApp = Application.Sw;

                if (firstBody == null) continue;

                SheetIndex++;
                string configName = $"Config_body_{SheetIndex}";
                string[] existingConfigs = (string[])swModel.GetConfigurationNames();

                // Create Derieved Configuration 
                bool warn = false;
                bool configExists = ((string[])swModel.GetConfigurationNames()).Contains(configName);

                if (!existingConfigs.Contains(configName))
                {
                    swModel.AddConfiguration3(
                        Name: configName,
                        Comment: $"Configuations for: {configName}",
                        AlternateName: "Configurations for each body",
                        (int)swConfigurationOptions2_e.swConfigOption_DontActivate
                    );

                    swModel.ShowConfiguration2(configName);
                    swModel.Extension.SelectAll();
                    swModel.HideSolidBody();

                    foreach (var body in part.Bodies)
                    {
                        var swBodyItem = ((ISwBody)body).Body;

                        bool isTarget = swBodyItem.Name == swBody.Name;
                        

                        if (isTarget)
                        {
                            var Model = swModel.Extension.SelectByID2(firstBody.Name, "SOLIDBODY", 0, 0, 0, true, 0, null, 0);
                            if (Model)
                            {
                                swModel.ShowSolidBody();

                            }

                        }

                        swModel.ClearSelection2(true);
                    }


                    }
                }


            }
        }
    }

            #endregion
            // Drawing Creation  Section 
       /*     #region

            string Template = @"C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2023\templates\RSRG A3.DRWDOT";
            var FirstSheet = Application.Documents.NewDrawing();



            int i = 0;
            while (i < SheetIndex)
            {
                AddSheet();
                i++;
            }
            #endregion

        }

       private void AddSheet()
        {
            var drawingDoc = Application.Documents.Active as ISwDrawing;
            var swModel = drawingDoc.Model as IModelDoc2;
            var swDrawing = swModel as SolidWorks.Interop.sldworks.DrawingDoc;

            string templatePath = @"C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2023\templates\RSRG A3.DRWDOT";

            // paper size enum (choose appropriate value)
            int paperSize = (int)swDwgPaperSizes_e.swDwgPaperA3size;
            // indicate custom template
            int templateType = (int)swDwgTemplates_e.swDwgTemplateCustom;

            bool ok = swDrawing.NewSheet4(
                "",
                paperSize,
                templateType,
                1.0,    // scale numerator
                1.0,    // scale denominator
                false,  // first angle projection? (false = third angle)
                templatePath,
                0, 0,    // width/height (0 = use template)
                "",     // property view name
                0, 0, 0, 0,// zone margins (left,right,top,bottom)
                0, 0     // zone rows, cols
            );

            if (!ok)
            {
                Application.ShowMessageBox($"Failed to create sheet  via SW API.");
            }
            else
            {
                // force update if needed
                swModel.ForceRebuild3(true);
            }
        }


  }
}
*/










