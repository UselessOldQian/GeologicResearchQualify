using System;
using System.Collections.Generic;
using System.Windows.Forms;
using ESRI.ArcGIS.esriSystem;

namespace Quality_Inspection_of_Overall_Planning_Results
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            //ESRI.ArcGIS.RuntimeManager.Bind(ESRI.ArcGIS.ProductCode.EngineOrDesktop);
            //#region
            //IAoInitialize m_AoInitialize = new AoInitializeClass();
            //esriLicenseStatus licenseStatus = esriLicenseStatus.esriLicenseUnavailable;
            //licenseStatus = m_AoInitialize.Initialize(esriLicenseProductCode.esriLicenseProductCodeAdvanced);
            //if (licenseStatus == esriLicenseStatus.esriLicenseNotInitialized)
            //{
            //    MessageBox.Show("esriLicenseProductCodeAdvanced许可！");
            //    Application.Exit();
            //}
            //#endregion
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            ESRI.ArcGIS.RuntimeManager.Bind(ESRI.ArcGIS.ProductCode.EngineOrDesktop);
            Application.Run(new SelectForm());
        }
    }
}