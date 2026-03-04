using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Pinnacle.UserControls
{
    public partial class CustomComboBox : UserControl
    {
        //TextBox textbox = new TextBox();
        //ToolStripDropDown _dropDown;
        ToolStrip toolStrip1 = new ToolStrip();
        internal ToolStripDropDownButton dropDownButton1;
        internal ToolStripDropDown dropDown;
        internal ToolStripButton buttonRed;
        internal ToolStripButton buttonBlue;
        internal ToolStripButton buttonYellow;
        //public class ToolStripDropDown : System.Windows.Forms.ToolStrip
        public CustomComboBox()
        {
            //InitializeComponent();


            //textbox.Location = new System.Drawing.Point(0, 0);
            //textbox.Size = new System.Drawing.Size(this.Width - 22, 20);
            //textbox.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
            //this.Controls.Add(textbox);
            //var button = new Button();
            //button.Location = new System.Drawing.Point(this.Width - 22, -1);
            //button.Size = new System.Drawing.Size(22, 22);
            //button.Text = "test";
            //button.Anchor = AnchorStyles.Right | AnchorStyles.Top;
            //button.BackgroundImage = global::Pinnacle.Properties.Resources.backward;
            //button.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            //button.Click += new System.EventHandler(this.Button_Click);
            //this.Controls.Add(button);
            //var dropDownControl = new DropDownControlTest();
            //var controlHost = new ToolStripControlHost(dropDownControl);
            //_dropDown = new ToolStripDropDown();
            //_dropDown.AutoSize = true;
            //_dropDown.Items.Add(controlHost);
            dropDownButton1 = new ToolStripDropDownButton();
            dropDown = new ToolStripDropDown();
            dropDownButton1.Text = "A";

            // Set the drop-down on the ToolStripDropDownButton.
            dropDownButton1.DropDown = dropDown;

            // Set the drop-down direction.
            dropDownButton1.DropDownDirection = ToolStripDropDownDirection.Left;

            // Do not show a drop-down arrow.
            dropDownButton1.ShowDropDownArrow = false;

            // Declare three buttons, set their foreground color and text, 
            // and add the buttons to the drop-down.
            buttonRed = new ToolStripButton();
            buttonRed.ForeColor = Color.Red;
            buttonRed.Text = "A";

            buttonBlue = new ToolStripButton();
            buttonBlue.ForeColor = Color.Blue;
            buttonBlue.Text = "A";

            buttonYellow = new ToolStripButton();
            buttonYellow.ForeColor = Color.Yellow;
            buttonYellow.Text = "A";

            buttonBlue.Click += new EventHandler(colorButtonsClick);
            buttonRed.Click += new EventHandler(colorButtonsClick);
            buttonYellow.Click += new EventHandler(colorButtonsClick);

            dropDown.Items.AddRange(new ToolStripItem[]
                { buttonRed, buttonBlue, buttonYellow });
            toolStrip1.Items.Add(dropDownButton1);
        }
        private void colorButtonsClick(object sender, EventArgs e)
        {
            ToolStripButton senderButton = (ToolStripButton)sender;
            this.ForeColor = senderButton.ForeColor;
        }
        //public class DropDownControlTest : UserControl
        //{

        //    public DropDownControlTest()
        //    {
        //        var listview = new ListView();
        //        listview.Location = new System.Drawing.Point(3, 1);
        //        listview.Size = new System.Drawing.Size(200, 100);
        //        listview.View = View.Details;
        //        listview.Columns.Add("Col 1", 100);
        //        listview.Columns.Add("Col 2", 100);
        //        this.Controls.Add(listview);
        //        //var button = new Button();
        //        //button.Location = new System.Drawing.Point(3, 105);
        //        //button.Text = "More...";
        //        //this.Controls.Add(button);

        //    }


        //}
        //void Button_Click(object sender, EventArgs e)
        //{

        //    _dropDown.Show(this, 0, this.Height);
        //}
    }
}
