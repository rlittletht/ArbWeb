using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ph
    {
    public partial class ArbDebug : Form
        {
        public AxSHDocVw.AxWebBrowser AxWeb
        {
            get { return m_axWebBrowser1; }
        }

        public ArbDebug()
            {
            InitializeComponent();
            }
        }
    }