using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using RhinoScript4;
using Rhino5x64;

namespace hola.RhinoConnect
{
    public class RhinoApplication
    {
        #region "Public Members"
        public IRhino5x64Application RhinoApp = null;
        public IRhinoScript RhinoScript = null;
        #endregion

        private string _progID = null;
        private bool _visible = false;

        public RhinoApplication(string id, bool IsVisible)
        {
            // widen scope
            _progID = id;
            _visible = IsVisible;

            //setup
            DoSetup();
        }

        /// <summary>
        /// Setup Rhino and RhinoScript
        /// </summary>
        private void DoSetup()
        {
            try
            {
                // Rhino Program ID
                string m_rhinoprogramID = _progID;
                IRhino5x64Application m_RhinoApp = null;

                // get Rhino application
                Type type = Type.GetTypeFromProgID(m_rhinoprogramID);
                dynamic rhino = Activator.CreateInstance(type); // Create Rhino instance
                m_RhinoApp = rhino as IRhino5x64Application;
                m_RhinoApp.Visible = Convert.ToInt16(_visible);  // 0 = hidden,  1 = visible

                RhinoApp = m_RhinoApp;  // Rhino Application

                IRhinoScript m_RhinoScript = null;
                m_RhinoScript = m_RhinoApp.GetScriptObject() as IRhinoScript;
                RhinoScript = m_RhinoScript;
            }
            catch { }
        }

    }
     
}
