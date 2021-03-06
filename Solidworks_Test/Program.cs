using System.Diagnostics;

namespace Solidworks_Test
{
    class Program
    {
        static void Main(string[] args)
        {
            SldWorks.SldWorks swApp = new SldWorks.SldWorks();
            int err = 0, warn = 0;

            Debug.Print(System.DateTime.Now.ToString());
            //SldWorks.ModelDoc2 swModel = swApp.LoadFile4("C:\\Users\\wgq\\Documents\\SldWorks\\test.IGS", "r", null, ref err);
            SldWorks.ModelDoc2 swModel = swApp.LoadFile4("C:\\Users\\wgq\\OneDrive\\Desktop\\ncon.stp", "r", null, ref err);
            //SldWorks.ModelDoc2 swModel = swApp.OpenDoc6(
            //    /* "C:\\Users\\wgq\\Documents\\SldWorks\\myNewPart.SLDPRT", */
            //    "C:\\Users\\Public\\Documents\\SOLIDWORKS\\SOLIDWORKS 2020\\samples\\tutorial\\tolanalyst\\offset\\top_plate.sldprt",
            //    /* "C:\\Users\\wgq\\Source\\Repos\\CSharpAndSolidWorks\\CSharpAndSolidWorks\\TemplateModel\\Measure.sldprt", */
            //    /* "C:\\Users\\Public\\Documents\\SOLIDWORKS\\SOLIDWORKS 2020\\samples\\tutorial\\routing-pipes\\ball valve with flanges.sldasm", */
            //    /* "C:\\Users\\Public\\Documents\\SOLIDWORKS\\SOLIDWORKS 2020\\samples\\tutorial\\advdrawings\\foodprocessor.slddrw", */
            //    (int)SwConst.swDocumentTypes_e.swDocPART,
            //    (int)SwConst.swOpenDocOptions_e/* .swOpenDocOptions_Silent */.swOpenDocOptions_ReadOnly,
            //    null, ref err, ref warn);
            if (swModel == null)
            {
                Debug.Print("--- !!! Open File Failed --> error --> " + err + " warning --> " + warn);
                swApp.ExitApp();
                swApp = null;
                return;
            }


            Debug.Print(System.DateTime.Now.ToString());
            if (swModel.GetCustomInfoValue("", "Project") != null)
                Debug.Print("--- 1. Info --> " + swModel.GetCustomInfoValue("", "Project"));


            Debug.Print(System.DateTime.Now.ToString());
            SldWorks.Configuration swConfig = default;
            if (swModel.GetConfigurationNames() != null)
                foreach (var name in swModel.GetConfigurationNames() as string[])
                {
                    swConfig = swModel.GetConfigurationByName(name);
                    var manager = swModel.Extension.CustomPropertyManager[name];
                    string code = manager.Get("Code");
                    var desc = manager.Get("Description");
                    Debug.Print("--- 2. Name of configuration ---> " + name + " Code = " + code + "Desc =" + desc);
                }


            System.Action<SldWorks.Feature, bool, bool> TraverseFeatures = null;
            TraverseFeatures = new System.Action<SldWorks.Feature, bool, bool>((SldWorks.Feature thisFeature, bool isTopLevel, bool isShowDimension) =>
            {
                SldWorks.Feature curFeature = thisFeature;

                while (curFeature != null)
                {
                    Debug.Print("--- Feature --> " + curFeature.Name);
                    if (isShowDimension == true)
                    {
                        SldWorks.DisplayDimension thisDisplayDimension = curFeature.GetFirstDisplayDimension();

                        while (thisDisplayDimension != null)
                        {
                            SldWorks.Dimension dimension = thisDisplayDimension.GetDimension();

                            Debug.Print($"--- Feature {curFeature.Name} Dimension --> " + thisDisplayDimension.GetNameForSelection() + " --> " + dimension.Value);

                            thisDisplayDimension = curFeature.GetNextDisplayDimension(thisDisplayDimension);
                        }
                    }

                    SldWorks.Feature subFeature = curFeature.GetFirstSubFeature();
                    while (subFeature != null)
                    {
                        TraverseFeatures(subFeature, false, false);
                        subFeature = subFeature.GetNextSubFeature();
                    }

                    if (isTopLevel)
                        curFeature = curFeature.GetNextFeature();
                    else
                        curFeature = null;
                }
            });

            Debug.Print(System.DateTime.Now.ToString());
            SldWorks.Feature swFeature = swModel.FirstFeature();
            if (swFeature != null)
            {
                Debug.Print("--- 3. Features ---");
                TraverseFeatures(swFeature, true, false);
            }


            swConfig = swModel.GetActiveConfiguration();
            if (swConfig != null)
            {
                System.Action<SldWorks.Component2, long> TraverseCompXform = null;
                TraverseCompXform = ((SldWorks.Component2 swComp, long nLevel) =>
                {
                    string sPadStr = "";
                    for (long i = 0; i < nLevel; i++)
                        sPadStr += "*";

                    SldWorks.MathTransform swCompXform = swComp.Transform2;
                    SldWorks.ModelDoc2 swModelOfComp = swComp.GetModelDoc2();

                    if (swCompXform != null)
                    {
                        Debug.Print("--- Pad String --> " + sPadStr + " " + swComp.Name2);
                        if (swComp.GetSelectByIDString() != "")
                            Debug.Print("--- Select ID --> " + swComp.GetSelectByIDString());

                        if (swModelOfComp != null)
                        {
                            if (swModelOfComp.GetType() == 1)
                                Debug.Print("--- Material --> " + (swModelOfComp as SldWorks.PartDoc).GetMaterialPropertyName2("", out string swMateDB)
                                    + " --- Database --> " + swMateDB);

                            Debug.Print("--- PartNum --> " + swModelOfComp.get_CustomInfo2(swComp.ReferencedConfiguration, "PartNum"));
                            Debug.Print("--- Name2 --> " + swComp.Name2);
                            Debug.Print("--- Name --> " + swModelOfComp.GetPathName());
                            Debug.Print("--- ConfigName --> " + swComp.ReferencedConfiguration);
                            Debug.Print("--- ComponentRef --> " + swComp.ComponentReference);
                        }
                    }

                    object[] vChild = swComp.GetChildren();
                    for (long i = 0; i <= (vChild.Length - 1); i++)
                    {
                        var swChildComp = vChild[i] as SldWorks.Component2;
                        TraverseCompXform(swChildComp, nLevel + 1);
                    }
                });

                Debug.Print(System.DateTime.Now.ToString());
                Debug.Print("--- 4. Active Configuration ---");
                SldWorks.Component2 swRootComp = swConfig.GetRootComponent();
                if (swRootComp != null)
                {
                    Debug.Print("--- 4.1. Root Component ---");
                    TraverseCompXform(swRootComp, 0);
                }
            }


            Debug.Print(System.DateTime.Now.ToString());
            SldWorks.DrawingDoc swDraw = swModel as SldWorks.DrawingDoc;
            if (swDraw != null)
            {
                Debug.Print("--- 5. Drawing Doc ---");
                var sheetNames = swDraw.GetSheetNames() as object[];

                string k3Name = "";
                foreach (var kName in sheetNames)
                {
                    Debug.Print("--- Sheet Name ---> " + kName);
                    if ((kName as string).Contains("k3"))
                        k3Name = kName as string;
                }

                bool bActSheet = swDraw.ActivateSheet(k3Name);

                SldWorks.Sheet drwSheet = swDraw.GetCurrentSheet();
                if (drwSheet != null)
                {
                    Debug.Print("--- 5.1. Current Sheet ---");
                    object[] views = drwSheet.GetViews() as object[];

                    if (views != null)
                    {
                        Debug.Print("--- 5.2. Get Views ---");
                        foreach (object vView in views)
                        {
                            var ss = vView as SldWorks.View;
                            Debug.Print("--- View --> " + ss.GetName2());
                        }
                    }
                }
            }
            swModel.ShowNamedView2("*Front", (int)SwConst.swStandardViews_e.swFrontView);


            if (swModel.Extension.SelectByID2("Sketch1", "SKETCH", 0, 0, 0, false, 0, null, 0))
                Debug.Print("--- select succeeded ---");
            SldWorks.SelectionMgr modelSel = swModel.SelectionManager;
            SldWorks.View actionView = modelSel.GetSelectedObject6(1, -1) as SldWorks.View;


            Debug.Print(System.DateTime.Now.ToString());
            var noteCount = 0;
            if (actionView != null)
            {
                Debug.Print("--- 6. Selected Object ---");
                noteCount = actionView.GetNoteCount();
            }

            if (noteCount > 0)
            {
                Debug.Print("--- Note Count --> " + noteCount.ToString());
                SldWorks.Note note = actionView.GetFirstNote();
                if (note != null)
                {
                    Debug.Print("--- 6.1. Notes ---");
                    Debug.Print("--- Components ---> ");
                    try
                    {
                        Debug.Print((((note.GetAnnotation() as SldWorks.Annotation)
                            .GetAttachedEntities3()[0] as SldWorks.Entity)
                            .GetComponent() as SldWorks.Component2)
                            .Name2);
                    }
                    catch
                    {
                        Debug.Print("--- !!! Error getting components ---");
                    };
                    Debug.Print("--- Note --> " + note.GetText());

                    //var leaderInfo = note.GetLeaderInfo();
                    for (int k = 0; k < noteCount - 1; k++)
                    {
                        note = note.GetNext() as SldWorks.Note;
                        Debug.Print("--- Note --> " + note.GetText());
                    }
                }
            }


            Debug.Print(System.DateTime.Now.ToString());
            Debug.Print("--- 7. Export ---");
            Debug.Print("--- Model Type --> " + swModel.GetType());
            SldWorks.ModelDocExtension swModExt = swModel.Extension;
            if (swModel.GetType() == (int)SwConst.swDocumentTypes_e.swDocPART || swModel.GetType() == (int)SwConst.swDocumentTypes_e.swDocASSEMBLY)
            {
                var setRes = swModel.Extension.SetUserPreferenceString(16, 0, "CustomerCS");
                swApp.SetUserPreferenceIntegerValue((int)SwConst.swUserPreferenceIntegerValue_e.swParasolidOutputVersion, (int)SwConst.swParasolidOutputVersion_e.swParasolidOutputVersion_270);
                swModExt.SaveAs(@"C:\Users\wgq\Documents\SldWorks\export.x_t", (int)SwConst.swSaveAsVersion_e.swSaveAsCurrentVersion, (int)SwConst.swSaveAsOptions_e.swSaveAsOptions_Silent, null, ref err, ref warn);
                Debug.Print("--- Part or Assembly Exported ---");
            }
            else if (swModel.GetType() == (int)SwConst.swDocTemplateTypes_e.swDocTemplateTypeDRAWING)
            {
                swApp.SetUserPreferenceIntegerValue((int)SwConst.swUserPreferenceIntegerValue_e.swDxfVersion, 2);
                swModel.SetUserPreferenceToggle(196, false);
                swModExt.SaveAs(@"C:\Users\wgq\Documents\SldWorks\export.dxf", (int)SwConst.swSaveAsVersion_e.swSaveAsCurrentVersion, (int)SwConst.swSaveAsOptions_e.swSaveAsOptions_Silent, null, ref err, ref warn);
                Debug.Print("--- Drawing Exported ---");
            }


            Debug.Print(System.DateTime.Now.ToString());
            if (swModExt.SelectByID2("Sketch1", "SKETCH", 0, 0, 0, false, 0, null, 0))
            {
                Debug.Print("--- 8. Traverse Sketch Segment ---");
                swFeature = modelSel.GetSelectedObject6(1, -1);
                swModel.EditSketch();
                SldWorks.Sketch sk = swFeature.GetSpecificFeature2();
                object[] vSketchSeg = sk.GetSketchSegments() as object[];

                SldWorks.SketchSegment swSketchSeg;
                double totalLength = 0;
                foreach (var tempSeg in vSketchSeg)
                {
                    swSketchSeg = tempSeg as SldWorks.SketchSegment;
                    if (swSketchSeg.GetType() != (int)SwConst.swSketchSegments_e.swSketchTEXT && swSketchSeg.ConstructionGeometry == false)
                        totalLength += swSketchSeg.GetLength();
                }

                swModel.EditSketch();
                Debug.Print("--- Total Length --> " + totalLength * 1000);
                //swApp.SendMsgToUser("Total Length: " + totalLength * 1000);
            }


            var OnSaveToStorage = new System.Func<int>(() =>
            {
                System.Runtime.InteropServices.ComTypes.IStream iStr = swModel.IGet3rdPartyStorage("Tool.Name", true);

                byte[] data = System.Text.Encoding.Unicode.GetBytes("Save String");
                iStr.Write(data, data.Length, System.IntPtr.Zero);

                swModel.IRelease3rdPartyStorage("Tool.Name");
                Debug.Print("--- Wrote in Callback ---");

                return 0;
            });


            Debug.Print(System.DateTime.Now.ToString());
            Debug.Print("--- 9. Write Third Party Data ---");
            switch (swModel.GetType())
            {
                case (int)SwConst.swDocumentTypes_e.swDocPART:
                    (swModel as SldWorks.PartDoc).SaveToStorageNotify += new SldWorks.DPartDocEvents_SaveToStorageNotifyEventHandler(OnSaveToStorage);
                    Debug.Print("--- Writing Part ---");
                    break;
                case (int)SwConst.swDocumentTypes_e.swDocASSEMBLY:
                    (swModel as SldWorks.AssemblyDoc).SaveToStorageNotify += new SldWorks.DAssemblyDocEvents_SaveToStorageNotifyEventHandler(OnSaveToStorage);
                    Debug.Print("--- Writing Assembly ---");
                    break;
                case (int)SwConst.swDocumentTypes_e.swDocDRAWING:
                    (swModel as SldWorks.DrawingDoc).SaveToStorageNotify += new SldWorks.DDrawingDocEvents_SaveToStorageNotifyEventHandler(OnSaveToStorage);
                    Debug.Print("--- Writing Drawing ---");
                    break;
            }
            //swModel.SetSaveFlag();
            if (!swModel.Save3((int)SwConst.swSaveAsOptions_e.swSaveAsOptions_Silent, ref err, ref warn))
                Debug.Print("--- !!! Failed to save model ---> error --> " + err + " warning --> " + warn);
            else
                Debug.Print("--- Save Completed ---" /* "--- Completed. Please save file. ---" */);

            Debug.Print("--- 9.1. Read Third Party Data ---");
            System.Runtime.InteropServices.ComTypes.IStream iStr = swModel.IGet3rdPartyStorage("Tool.Name", false);
            if (iStr != null)
            {
                Debug.Print("--- Got Data ---");

                iStr.Stat(out System.Runtime.InteropServices.ComTypes.STATSTG statstg, 1 /* STATSFLAG_NONAME */);

                long length = statstg.cbSize;
                byte[] data = new byte[length];
                System.IntPtr bytesRead = default(System.IntPtr);

                iStr.Read(data, (int)length, bytesRead);
                string strData = System.Text.Encoding.Unicode.GetString(data);
                Debug.Print("--- Data --> " + strData);
            }
            swModel.IRelease3rdPartyStorage("Tool.Name");


            SldWorks.Frame swFrame = swApp.Frame();
            swFrame.SetStatusBarText("Status Bar Text --> ");
            swApp.GetUserProgressBar(out SldWorks.UserProgressBar userProgressBar);
            userProgressBar.Start(0, 100, "Status");

            int position = 0;
            for (int i = 0; i <= 100; i++)
            {
                position = i * 10;

                if (position == 100)
                {
                    position = 0;
                    break;
                }

                userProgressBar.UpdateProgress(position);
                userProgressBar.UpdateTitle("Progress --> " + position);
            }
            userProgressBar.End();


            /* Advanced Component Selection part is intentionally skipped */


            Debug.Print(System.DateTime.Now.ToString());
            if (swModel.GetType() == (int)SwConst.swDocumentTypes_e.swDocPART)
            {
                Debug.Print("--- 11. Bounding Box on Part Doc ---");
                swFeature = (swModel as SldWorks.PartDoc).FeatureByName("Bounding Box");
                if (swFeature == null)
                {
                    swFeature = swModel.FeatureManager.InsertGlobalBoundingBox((int)SwConst.swGlobalBoundingBoxFitOptions_e.swBoundingBoxType_BestFit, true, false, out int status);
                    Debug.Print("--- status --> " + status);
                }

                swModel.SetUserPreferenceToggle((int)SwConst.swUserPreferenceToggle_e.swViewDispGlobalBBox, true);

                SldWorks.Configuration configuration = swModel.GetActiveConfiguration();
                SldWorks.CustomPropertyManager manager2 = swModel.Extension.get_CustomPropertyManager(configuration.Name);

                manager2.Get3("Total Bounding Box Length", true, out string str, out string str2);
                manager2.Get3("Total Bounding Box Width", true, out str, out string str3);
                manager2.Get3("Total Bounding Box Thickness", true, out str, out string str4);

                Debug.Print("--- Bounding Box Size --> Length --> " + str2 + " Width --> " + str3 + " Thickness --> " + str4);
            }


            Debug.Print(System.DateTime.Now.ToString());
            swModExt = swModel.Extension;
            SldWorks.Measure swMeasure = swModExt.CreateMeasure();
            swMeasure.ArcOption = 0;
            bool stat = swMeasure.Calculate(null);
            if (stat) { Debug.Print("--- 12. Measure Distance --> " + (swMeasure.Distance * 1000)); }
            else { Debug.Print("--- 12. Measure Distance --> status --> " + stat); }


            swModExt.IncludeMassPropertiesOfHiddenBodies = false;
            int massStatus = 0;
            double[] massProperties = swModExt.GetMassProperties(1, ref massStatus);
            if (massProperties != null)
            {
                Debug.Print("--- 13. Mass Properties -->");
                Debug.Print("--- Center of Mass X --> " + massProperties[0]);
                Debug.Print("--- Center of Mass Y --> " + massProperties[1]);
                Debug.Print("--- Center of Mass Z --> " + massProperties[2]);
                Debug.Print("--- Volume --> " + massProperties[3]);
                Debug.Print("--- Area --> " + massProperties[4]);
                Debug.Print("--- Mass --> " + massProperties[5]);
                Debug.Print("--- Moment XX --> " + massProperties[6]);
                Debug.Print("--- Moment YY --> " + massProperties[7]);
                Debug.Print("--- Moment ZZ --> " + massProperties[8]);
                Debug.Print("--- Moment XY --> " + massProperties[9]);
                Debug.Print("--- Moment ZX --> " + massProperties[10]);
                Debug.Print("--- Moment YZ --> " + massProperties[11]);
            }


            /* Bom List functionalities (mostly) merged into TraverseCompXform() */


            Debug.Print(System.DateTime.Now.ToString());
            for (int i = 1; i < modelSel.GetSelectedObjectCount(); i++)
            {
                SldWorks.Face2 face2 = modelSel.GetSelectedObject6(1, -1);
                object[] vFaceProp = swModel.MaterialPropertyValues;

                object[] vProps = face2.GetMaterialPropertyValues2(1, null);
                vProps[0] = 1; // red
                vProps[1] = 0;
                vProps[2] = 0;
                vProps[3] = vFaceProp[3];
                vProps[4] = vFaceProp[4];
                vProps[5] = vFaceProp[5];
                vProps[6] = vFaceProp[6];
                vProps[7] = vFaceProp[7];
                vProps[8] = vFaceProp[8];
                face2.SetMaterialPropertyValues2(vProps, 1, null);

                vProps = null;
                vFaceProp = null;
            }
            swModel.ClearSelection2(true);
            Debug.Print("--- 14. Set Color (Method 1) Completed ---");

            double[] matPropVals = swModel.MaterialPropertyValues;
            System.Random rnd = new System.Random();

            var tempC = System.IO.Path.GetFileNameWithoutExtension(swModel.GetPathName()).Contains("m1") ?
                System.Drawing.Color.Red :
                System.Drawing.Color.FromArgb(
                    rnd.Next(0, 255),
                    rnd.Next(0, 255),
                    rnd.Next(0, 255));
            matPropVals[0] = System.Convert.ToDouble(tempC.R) / 255;
            matPropVals[1] = System.Convert.ToDouble(tempC.G) / 255;
            matPropVals[2] = System.Convert.ToDouble(tempC.B) / 255;
            swModel.MaterialPropertyValues = matPropVals;

            swModel.WindowRedraw();
            Debug.Print("--- 14.1. Set Color (Method 2) Completed ---");


            Debug.Print(System.DateTime.Now.ToString());
            if (actionView != null)
            {
                SldWorks.ModelDoc2 viewModel = actionView.ReferencedDocument;
                Debug.Print("--- 15. Get Drawing Model --> Path Name --> " + viewModel.GetPathName());

                string stepName = System.IO.Path.GetFileNameWithoutExtension(viewModel.GetPathName());
                if (viewModel.Extension.SaveAs(@"C:\Users\wgq\Documents\SldWorks\{stepName}.step", (int)SwConst.swSaveAsVersion_e.swSaveAsCurrentVersion,
                    (int)SwConst.swSaveAsOptions_e.swSaveAsOptions_Silent, null, ref err, ref warn))
                    Debug.Print("--- 15.1. Export Drawing Model Done ---");
                else
                    Debug.Print("--- !!! Export Drawing Model Failed --> error --> " + err + " warning --> " + warn);
            }


            Debug.Print(System.DateTime.Now.ToString());
            if (swDraw != null)
            {
                Debug.Print("--- 99. Views ---");
                SldWorks.View swView = swDraw.GetFirstView(), swBaseView;

                while (swView != null)
                {
                    swBaseView = swView.GetBaseView();
                    Debug.Print("--- View ---> " + swView.Name);

                    if (swBaseView != null)
                        Debug.Print("--- Base View --> ", swBaseView.Name);

                    swView = swView.GetNextView();
                }
            }

            swApp.ExitApp();
            swApp = null;
        }
    }
}
