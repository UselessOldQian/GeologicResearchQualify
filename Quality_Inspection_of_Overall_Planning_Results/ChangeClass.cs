using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Quality_Inspection_of_Overall_Planning_Results
{
    public class ChangeEventArgs : EventArgs
    {
        public ChangeEventArgs()
        { }

        private string _field_value = "";
        public string field_value
        {
            get { return _field_value; }
            set { _field_value = value; }
        }
    }
}
