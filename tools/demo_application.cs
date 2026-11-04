using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

internal static class DemoApplication
{
    [STAThread]
    private static void Main()
    {
        Application.EnableVisualStyles();
        string marker = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "launched.txt");
        using (Form form = new Form())
        using (Timer timer = new Timer())
        {
            form.Text = "TabEx Demo App";
            form.StartPosition = FormStartPosition.Manual;
            form.SetBounds(650, 260, 360, 160);
            form.TopMost = true;
            form.Controls.Add(new Label {
                Text = "Launch successful!\r\nThis demo closes automatically.",
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            });
            timer.Interval = 2800;
            timer.Tick += delegate { form.Close(); };
            form.Shown += delegate {
                File.WriteAllText(marker, Process.GetCurrentProcess().Id.ToString());
                timer.Start();
            };
            Application.Run(form);
        }
    }
}
