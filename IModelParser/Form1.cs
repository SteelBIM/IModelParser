using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Bentley.IModel.Core;
using Bentley.IModel.Core.BusinessData;
using Bentley.IModel.Core.MetaData;

namespace IModelParser
{
    public partial class Form1 : Form
    {
        private String filepath = null;
        private Dictionary<string, Dictionary<string, string>> result = null;

        public Form1()
        {
            InitializeComponent();
        }

        private void OpenIModel(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            if(fileDialog.ShowDialog() == DialogResult.OK)
            {
                filepath = fileDialog.FileName;
                if(filepath.EndsWith(".i.dgn", true, null))
                {
                    parseIModel();
                    drawTable();
                }
                else
                {
                    string[] exts = filepath.Split('.');
                    int length = exts.Length;
                    string ext = null;
                    if(length >= 2)
                    {
                        for (int i = 1; i < length; i++)
                        {
                            ext += exts[i];
                        }
                    }
                    MessageBox.Show("错误文件的后缀为:" + ext);
                }
            }
        }

        private void parseIModel()
        {
            result = new Dictionary<string, Dictionary<string, string>>();
            
            IModel imodel = IModel.Open(filepath);
            IEnumerable<IModelElement> elements = imodel.Elements;
            foreach(IModelElement element in elements)
            {

                string eID = element.Element.ElementId.ToString();
                string eName = element.Element.TypeName;
                string eKey = eName + eID;
                Dictionary<string, string> props = new Dictionary<string, string>();
                foreach(Dynamics o in element.Objects)
                {
                    foreach(Property p in o.Class.Properties)
                    {
                        if(!o.ECInstance[p.Name].IsNull)
                        {
                            props.Add(p.Name, o.ECInstance[p.Name].NativeValue.ToString());
                        }
                        else
                        {
                            props.Add(p.Name, "null");
                        }
                    }
                }
                result.Add(eKey, props);
            }
        }
        
        private void drawTable()
        {
            Dictionary<string, Dictionary<string, string>>.KeyCollection keys = result.Keys;
            foreach(string key in keys)
            {
                this.listBox1.Items.Add(key);
            }

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.listView1.Items.Clear();
            if(this.listBox1.SelectedItems.Count > 0)
            {
                string selectedItem = this.listBox1.SelectedItem.ToString();
                Dictionary<string, string> props = result[selectedItem];
                if(props.Keys.Count > 0)
                {
                    foreach(string propName in props.Keys)
                    {
                        ListViewItem viewItem = new ListViewItem();
                        viewItem.Text = propName;
                        viewItem.SubItems.Add(props[propName]);
                        this.listView1.Items.Add(viewItem);
                    }
                }
            }
        }

        private void outputTxt(object sender, EventArgs e)
        {
            if(result == null)
            {
                MessageBox.Show("尚未解析imodel");
            }
            else
            {
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "txt(*.txt)|*.txt";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    string filepath = sfd.FileName.ToString();
                    FileOutput fo = new FileOutput();
                    fo.outputTxt(filepath, result);
                }
            }

        }

        private void outputExcel(object sender, EventArgs e)
        {
            if(result == null)
            {
                MessageBox.Show("尚未解析imodel");
            }
            else
            {
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "Excel(*.xlsx)|*.xlsx";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    string filepath = sfd.FileName.ToString();
                    FileOutput fo = new FileOutput();
                    fo.outputExcel(filepath, result);
                }
            }
        }

        private void exit(object sender, EventArgs e)
        {
            this.Close();
        }

        private void regex(object sender, EventArgs e)
        {
            MatchForm mf = new MatchForm();
            mf.ShowDialog();
        }

    }
}
