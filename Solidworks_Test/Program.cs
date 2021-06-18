using System.Diagnostics;

namespace Solidworks_Test
{
    class Program
    {
        public static void ShowDimensionForFeature(SldWorks.Feature feature)
        {
            var thisDisplayDimension = feature.GetFirstDisplayDimension() as SldWorks.DisplayDimension;

            while (thisDisplayDimension != null)
            {
                var dimension = thisDisplayDimension.GetDimension() as SldWorks.Dimension;

                Debug.Print($"--- Feature {feature.Name} Dimension --> " + thisDisplayDimension.GetNameForSelection() + " --> " + dimension.Value);

                thisDisplayDimension = feature.GetNextDisplayDimension(thisDisplayDimension) as SldWorks.DisplayDimension;
            }
        }

        public static void TraverseFeatures(SldWorks.Feature thisFeature, bool isTopLevel, bool isShowDimension = false)
        {
            SldWorks.Feature curFeature = thisFeature;

            while (curFeature != null)
            {
                Debug.Print("--- Feature --> " + curFeature.Name);
                if (isShowDimension == true) ShowDimensionForFeature(curFeature);

                SldWorks.Feature subFeature = curFeature.GetFirstSubFeature() as SldWorks.Feature;
                while (subFeature != null)
                {
                    TraverseFeatures(subFeature, false);
                    subFeature = subFeature.GetNextSubFeature() as SldWorks.Feature;
                }

                if (isTopLevel)
                {
                    curFeature = curFeature.GetNextFeature() as SldWorks.Feature;
                }
                else
                {
                    curFeature = null;
                }
            }
        }

        public static void TraverseCompXform(SldWorks.Component2 swComp, long nLevel, bool setcolor = false)
        {
            object[] vChild;
            SldWorks.Component2 swChildComp;
            string sPadStr = "";
            SldWorks.MathTransform swCompXform;

            for (long i = 0; i < nLevel; i++)
            {
                sPadStr = sPadStr + "*";
            }
            swCompXform = swComp.Transform2;
            SldWorks.ModelDoc2 swModel = swComp.GetModelDoc2() as SldWorks.ModelDoc2;

            if (swCompXform != null)
            {
                try
                {
                    Debug.Print("--- Pad String --> " + sPadStr + " " + swComp.Name2);

                    if (swComp.GetSelectByIDString() == "")
                    {
                        Debug.Print("--- Select ID --> " + swComp.GetSelectByIDString());
                    }
                }
                catch
                {
                    Debug.Print("Exception catched");
                }

                if (swModel != null)
                {
                    Debug.Print("--- PartNum --> " + swModel.get_CustomInfo2(swComp.ReferencedConfiguration, "PartNum"));
                    Debug.Print("--- Name2 --> " + swComp.Name2);
                    Debug.Print("--- Name --> " + swModel.GetPathName());
                    Debug.Print("--- ConfigName --> " + swComp.ReferencedConfiguration);
                    Debug.Print("--- ComponentRef --> " + swComp.ComponentReference);

                    if (setcolor == true)
                    {
                        double[] matPropVals = swModel.MaterialPropertyValues as double[];
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
                    }
                }
            }

            vChild = swComp.GetChildren() as object[];
            for (long i = 0; i <= (vChild.Length - 1); i++)
            {
                swChildComp = vChild[i] as SldWorks.Component2;
                TraverseCompXform(swChildComp, nLevel + 1, setcolor);
            }
        }

        static void Main(string[] args)
        {
            SldWorks.SldWorks swApp = new SldWorks.SldWorks();
            int err = 0, warn = 0;
            SldWorks.ModelDoc2 swModel = swApp.LoadFile4("C:\\Users\\wgq\\OneDrive\\Desktop\\test.IGS", "r", null, ref err);
            /*SldWorks.ModelDoc2 swModel = swApp.OpenDoc6("C:\\Users\\wgq\\myNewPart.SLDPRT",
                (int)SwConst.swDocumentTypes_e.swDocPART,
                (int)SwConst.swOpenDocOptions_e.swOpenDocOptions_ReadOnly,
                null, ref err, ref warn);*/
            if (swModel.GetCustomInfoValue("", "Project") != null)
                Debug.Print("--- 1. Info --> " + swModel.GetCustomInfoValue("", "Project"));

            SldWorks.Configuration swConfig = default(SldWorks.Configuration);
            if (swModel.GetConfigurationNames() != null)
                foreach (var name in swModel.GetConfigurationNames() as string[])
                {
                    swConfig = swModel.GetConfigurationByName(name) as SldWorks.Configuration;
                    var manager = swModel.Extension.CustomPropertyManager[name];
                    string code = manager.Get("Code");
                    var desc = manager.Get("Description");
                    Debug.Print("--- 2. Name of configuration ---> " + name + " Code = " + code + "Desc =" + desc);
                }

            SldWorks.Feature swFeature = swModel.FirstFeature() as SldWorks.Feature;
            if (swFeature != null)
            {
                Debug.Print("--- 3. Features ---");
                TraverseFeatures(swFeature, true);
            }

            swConfig = swModel.GetActiveConfiguration() as SldWorks.Configuration;
            if (swConfig != null)
            {
                Debug.Print("--- 4. Active Configuration ---");
                SldWorks.Component2 swRootComp = swConfig.GetRootComponent() as SldWorks.Component2;
                if (swRootComp != null)
                {
                    Debug.Print("--- 4.1. Root Component ---");
                    TraverseCompXform(swRootComp, 0);
                }
            }

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

                SldWorks.Sheet drwSheet = swDraw.GetCurrentSheet() as SldWorks.Sheet;
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

            SldWorks.SelectionMgr modelSel = swModel.ISelectionManager;
            SldWorks.View actionView = modelSel.GetSelectedObject5(1) as SldWorks.View;

            var noteCount = 0;
            if (actionView != null)
            {
                Debug.Print("--- 6. Selected Object ---");
                noteCount = actionView.GetNoteCount();
            }

            if (noteCount > 0)
            {
                Debug.Print("--- Note Count --> " + noteCount.ToString());
                SldWorks.Note note = actionView.GetFirstNote() as SldWorks.Note;
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
                    catch { };
                    Debug.Print("--- Note --> " + note.GetText());

                    //var leaderInfo = note.GetLeaderInfo();
                    for (int k = 0; k < noteCount - 1; k++)
                    {
                        note = note.GetNext() as SldWorks.Note;
                        Debug.Print("--- Note --> " + note.GetText());
                    }
                }
            }

            Debug.Print("--- 7. Export ---");
            Debug.Print("--- Model Type --> " + swModel.GetType());
            SldWorks.ModelDocExtension swModExt = swModel.Extension;
            if (swModel.GetType() == (int)SwConst.swDocumentTypes_e.swDocPART || swModel.GetType() == (int)SwConst.swDocumentTypes_e.swDocASSEMBLY)
            {
                var setRes = swModel.Extension.SetUserPreferenceString(16, 0, "CustomerCS");
                swApp.SetUserPreferenceIntegerValue((int)SwConst.swUserPreferenceIntegerValue_e.swParasolidOutputVersion, (int)SwConst.swParasolidOutputVersion_e.swParasolidOutputVersion_270);
                swModExt.SaveAs(@"C:\Users\wgq\export.x_t", (int)SwConst.swSaveAsVersion_e.swSaveAsCurrentVersion, (int)SwConst.swSaveAsOptions_e.swSaveAsOptions_Silent, null, ref err, ref warn);
            }
            else if (swModel.GetType() == (int)SwConst.swDocTemplateTypes_e.swDocTemplateTypeDRAWING)
            {
                swApp.SetUserPreferenceIntegerValue((int)SwConst.swUserPreferenceIntegerValue_e.swDxfVersion, 2);
                swModel.SetUserPreferenceToggle(196, false);
                swModExt.SaveAs(@"C:\Users\wgq\export.dxf", (int)SwConst.swSaveAsVersion_e.swSaveAsCurrentVersion, (int)SwConst.swSaveAsOptions_e.swSaveAsOptions_Silent, null, ref err, ref warn);
            }

            if (swDraw != null)
            {
                Debug.Print("--- 99. Views ---");
                SldWorks.View swView = swDraw.GetFirstView(), swBaseView;

                while (swView != null)
                {
                    swBaseView = swView.GetBaseView();
                    Debug.Print("--- View ---> " + swView.Name);

                    if (swBaseView != null)
                    {
                        Debug.Print("--- Base View --> ", swBaseView.Name);
                    }

                    swView = swView.GetNextView();
                }
            }

            swApp.ExitApp();
            swApp = null;
        }
    }
}
