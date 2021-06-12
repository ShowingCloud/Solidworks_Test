using System.Diagnostics;

namespace Solidworks_Test
{
    class Program
    {
        public static void ShowDimensionForFeature(SldWorks.Feature feature)
        {
            var thisDisplayDimension = (SldWorks.DisplayDimension)feature.GetFirstDisplayDimension();

            while (thisDisplayDimension != null)
            {
                var dimension = (SldWorks.Dimension)thisDisplayDimension.GetDimension();

                Debug.Print($"--- Feature {feature.Name} Dimension --> " + thisDisplayDimension.GetNameForSelection() + " --> " + dimension.Value);

                thisDisplayDimension = (SldWorks.DisplayDimension)feature.GetNextDisplayDimension(thisDisplayDimension);
            }
        }

        public static void TraverseFeatures(SldWorks.Feature thisFeature, bool isTopLevel, bool isShowDimension = false)
        {
            SldWorks.Feature curFeature = default(SldWorks.Feature);
            curFeature = thisFeature;

            while (curFeature != null)
            {
                Debug.Print("--- Feature --> " + curFeature.Name);
                if (isShowDimension == true) ShowDimensionForFeature(curFeature);

                SldWorks.Feature subFeature = default(SldWorks.Feature);
                subFeature = (SldWorks.Feature)curFeature.GetFirstSubFeature();

                while (subFeature != null)
                {
                    TraverseFeatures(subFeature, false);
                    SldWorks.Feature nextSubFeature = default(SldWorks.Feature);
                    nextSubFeature = (SldWorks.Feature)subFeature.GetNextSubFeature();
                    subFeature = nextSubFeature;
                    nextSubFeature = null;
                }

                subFeature = null;

                SldWorks.Feature nextFeature = default(SldWorks.Feature);

                if (isTopLevel)
                {
                    nextFeature = (SldWorks.Feature)curFeature.GetNextFeature();
                }
                else
                {
                    nextFeature = null;
                }

                curFeature = nextFeature;
                nextFeature = null;
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
            SldWorks.ModelDoc2 swModel = (SldWorks.ModelDoc2)swComp.GetModelDoc2();

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
                        double[] matPropVals = (double[])swModel.MaterialPropertyValues;
                        System.Random rnd = new System.Random();

                        var tempC = default(System.Drawing.Color);
                        if (System.IO.Path.GetFileNameWithoutExtension(swModel.GetPathName()).Contains("m1"))
                        {
                            tempC = System.Drawing.Color.Red;
                        }
                        else
                        {
                            tempC = System.Drawing.Color.FromArgb(
                                rnd.Next(0, 255),
                                rnd.Next(0, 255),
                                rnd.Next(0, 255));
                        }
                        matPropVals[0] = System.Convert.ToDouble(tempC.R) / 255;
                        matPropVals[1] = System.Convert.ToDouble(tempC.G) / 255;
                        matPropVals[2] = System.Convert.ToDouble(tempC.B) / 255;
                        swModel.MaterialPropertyValues = matPropVals;

                        swModel.WindowRedraw();
                    }
                }
            }

            vChild = (object[])swComp.GetChildren();
            for (long i = 0; i <= (vChild.Length - 1); i++)
            {
                swChildComp = (SldWorks.Component2)vChild[i];
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
            foreach (var name in (string[])swModel.GetConfigurationNames())
            {
                swConfig = (SldWorks.Configuration)swModel.GetConfigurationByName(name);
                var manager = swModel.Extension.CustomPropertyManager[name];
                string code = manager.Get("Code");
                var desc = manager.Get("Description");
                Debug.Print("   Name of configuration ---> " + name + " Code = " + code + "Desc =" + desc);
            }

            SldWorks.Feature swFeature = (SldWorks.Feature)swModel.FirstFeature();
            TraverseFeatures(swFeature, true);

            swConfig = (SldWorks.Configuration)swModel.GetActiveConfiguration();
            SldWorks.Component2 swRootComp = (SldWorks.Component2)swConfig.GetRootComponent();
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
