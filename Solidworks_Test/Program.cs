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
            int err = 0;
            SldWorks.ModelDoc2 swModel = swApp.LoadFile4("C:\\Users\\wgq\\OneDrive\\Desktop\\test.IGS", "r", null, ref err);
            Debug.Print("Info: " + swModel.GetCustomInfoValue("", "Project"));

            SldWorks.Configuration swConfig = default(SldWorks.Configuration);
            foreach (var name in swModel.GetConfigurationNames() as string[])
            {
                swConfig = swModel.GetConfigurationByName(name) as SldWorks.Configuration;
                var manager = swModel.Extension.CustomPropertyManager[name];
                string code = manager.Get("Code");
                var desc = manager.Get("Description");
                Debug.Print("   Name of configuration ---> " + name + " Code = " + code + "Desc =" + desc);
            }

            SldWorks.Feature swFeature = swModel.FirstFeature() as SldWorks.Feature;
            TraverseFeatures(swFeature, true);

            swConfig = swModel.GetActiveConfiguration() as SldWorks.Configuration;
            SldWorks.Component2 swRootComp = swConfig.GetRootComponent() as SldWorks.Component2;
            TraverseCompXform(swRootComp, 0);

            SldWorks.DrawingDoc swDraw = swModel as SldWorks.DrawingDoc;
            SldWorks.View swView = swDraw.GetFirstView(), swBaseView;

            Debug.Print("File = " + swModel.GetPathName());

            while(swView != null)
            {
                swBaseView = swView.GetBaseView();
                Debug.Print("  " + swView.Name);

                if(swBaseView != null)
                {
                    Debug.Print("  --> ", swBaseView.Name);
                }

                swView = swView.GetNextView();
            }

            swApp.ExitApp();
            swApp = null;
        }
    }
}
