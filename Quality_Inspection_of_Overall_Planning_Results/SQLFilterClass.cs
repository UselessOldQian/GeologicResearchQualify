using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Quality_Inspection_of_Overall_Planning_Results
{
    public class SQLFileterEventArgs : EventArgs
    {
        public SQLFileterEventArgs()
        { }

        private int _layerIndex = -1;
        public int LayerIndex
        {
            get { return _layerIndex; }
            set { _layerIndex = value; }
        }

        private string _SQL = "";
        public string SQL
        {
            get { return _SQL; }
            set { _SQL = value; }
        }
        private string _SQL_2 = "";
        public string SQL_2
        {
            get { return _SQL_2; }
            set { _SQL_2 = value; }
        }
    }
}
