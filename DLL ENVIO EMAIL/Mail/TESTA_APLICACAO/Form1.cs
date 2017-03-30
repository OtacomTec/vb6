using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using clsMail;

namespace TESTA_APLICACAO
{
    public partial class Form1 : Form
    {
        string ABC;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
            clsMail.SmtpEnvio  a = new clsMail.SmtpEnvio();

            ABC = a.Envia("smtp.onlytechsolutions.com.br", "587", "marcos@onlytechsolutions.com.br", "TESTE", "marcos_onlytech@hotmail.com", "MARCOS É LINDO", "TESTE", "marcos@onlytechsolutions.com.br", "mar123***",false ); 

        }
    }
}
