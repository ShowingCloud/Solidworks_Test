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

                Debug.Print($"---Feature {feature.Name} Dimension -->" + thisDisplayDimension.GetNameForSelection() + "-->" + dimension.Value);

                thisDisplayDimension = (SldWorks.DisplayDimension)feature.GetNextDisplayDimension(thisDisplayDimension);
            }
        }
        public static void TraverseFeatures(SldWorks.Feature thisFeature, bool isTopLevel, bool isShowDimension = false)
        {
            SldWorks.Feature curFeature = default(SldWorks.Feature);
            curFeature = thisFeature;

            while (curFeature != null)
            {
                Debug.Print(curFeature.Name);
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

        static void Main(string[] args)
        {
            SldWorks.SldWorks swApp = new SldWorks.SldWorks();
            int err = 0;
            SldWorks.ModelDoc2 swModel = swApp.LoadFile4("C:\\Users\\wgq\\OneDrive\\Desktop\\test.IGS", "r", null, ref err);
            Debug.Print("Info: " + swModel.GetCustomInfoValue("", "Project"));
            foreach(var name in (string[])swModel.GetConfigurationNames())
            {
                var swConfig = (SldWorks.Configuration)swModel.GetConfigurationByName(name);
                var manager = swModel.Extension.CustomPropertyManager[name];
                string code = manager.Get("Code");
                Debug.Print("   Name of configuration ---> " + name + " Desc.=" + code);
            }

            SldWorks.Feature swFeature = (SldWorks.Feature)swModel.FirstFeature();
            TraverseFeatures(swFeature, true);


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
