using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ShortCuts
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            this.Hide();
            notifyIcon1.Visible = true;
            if (!Directory.Exists("shortcuts")) Directory.CreateDirectory("shortcuts");
        }





        private void contextMenuStrip1_Closed(object sender, ToolStripDropDownClosedEventArgs e)
        {
            contextMenuStrip1 = new ContextMenuStrip();
            contextMenuStrip1.Opened += contextMenuStrip1_Opened;
            contextMenuStrip1.Items.Add("Exit").Name="Exit";
            contextMenuStrip1.Items.Add("Reload").Name = "Reload";
            contextMenuStrip1.Items.Add("Open").Name = "Open";
        }

        private void contextMenuStrip1_Opened(object sender, EventArgs e)
        {
            contextMenuStrip1.Items.Clear();
            string[] files = Directory.GetFiles("shortcuts");
            for (int i = 0; i < files.Length; i++)
            {
                string ext = Path.GetExtension(files[i].ToString());//拡張子

                contextMenuStrip1.Items.Add(files[i]);
                contextMenuStrip1.Items[i].Name = files[i];
                contextMenuStrip1.Items[i].Text = files[i].Replace("shortcuts\\","");//いらない奴消す
                contextMenuStrip1.Items[i].Click += OnPressFile;

                //画像の設定
                if (ext == ".lnk")
                {
                    string path = "";
                    try
                    {
                        IWshRuntimeLibrary.WshShell shell = new IWshRuntimeLibrary.WshShell();
                        IWshRuntimeLibrary.IWshShortcut shortcut = (IWshRuntimeLibrary.IWshShortcut)shell.CreateShortcut(files[i].ToString());
                        string targetPath = shortcut.TargetPath.ToString();
                        path = files[i];
                        Icon appIcon =
                          System.Drawing.Icon.ExtractAssociatedIcon(shortcut.TargetPath);
                        contextMenuStrip1.Items[i].Image = appIcon.ToBitmap();
                    }//ショートカット先のパスを取得しビットマップを取得
                    catch (Exception ex)
                    {
                        try
                        {
                            IWshRuntimeLibrary.WshShell shell = new IWshRuntimeLibrary.WshShell();
                            IWshRuntimeLibrary.IWshShortcut shortcut = (IWshRuntimeLibrary.IWshShortcut)shell.CreateShortcut(files[i].ToString());
                            string targetPath = shortcut.TargetPath.ToString();

                            path = files[i];
                            Icon appIcon =
                              System.Drawing.Icon.ExtractAssociatedIcon(shortcut.TargetPath.Replace("C:\\Program Files (x86)\\", "C:\\Program Files\\"));
                            contextMenuStrip1.Items[i].Image = appIcon.ToBitmap();
                        }//(x86)を消してビットマップ取得。
                        catch (Exception ee)
                        {
                            try
                            {
                                path = files[i];
                                Icon appIcon =
                                  System.Drawing.Icon.ExtractAssociatedIcon(path);
                                contextMenuStrip1.Items[i].Image = appIcon.ToBitmap();
                            }//ショートカットからビットマップ取得
                            catch (Exception ee2)
                            {

                            }
                        }
                    }
                }//ショートカットファイルだったら
                else
                {
                    string path = files[i];
                    Icon appIcon =
                      System.Drawing.Icon.ExtractAssociatedIcon(path);
                    contextMenuStrip1.Items[i].Image = appIcon.ToBitmap();
                }
                //画像の設定終わり
            }

            {
                contextMenuStrip1.Items.Add("Open").Name = "Open";
                contextMenuStrip1.Items.Find("Open", false)[0].Text = "Open";
                contextMenuStrip1.Items.Find("Open", false)[0].Click += OnPressOpen;


                contextMenuStrip1.Items.Add("Reload").Name = "Reload";
                contextMenuStrip1.Items.Find("Reload", false)[0].Text = "Reload";
                contextMenuStrip1.Items.Find("Reload", false)[0].Click += OnPressReload;

                contextMenuStrip1.Items.Add("Exit").Name = "Exit";
                contextMenuStrip1.Items.Find("Exit", false)[0].Text = "Exit";
                contextMenuStrip1.Items.Find("Exit", false)[0].Click += OnPressExit;
            }//デフォルトのアイテム

            foreach (ToolStripItem o in contextMenuStrip1.Items)
            {
                Console.WriteLine(o.Name);
            }
            for(int i = 0; i < contextMenuStrip1.Items.Count; i++)
            {
                contextMenuStrip1.Items[i].ForeColor = Color.Black;
                contextMenuStrip1.Items[i].BackColor = Color.FromArgb(255, 255, 255, 255);
            }//アイテムの見た目設定。
             
        }
        private void OnPressExit(object sender,EventArgs e)=> this.Close();
        private void OnPressFile(object sender,EventArgs e)
        {
            string tag = "shortcuts\\"+sender.ToString();
            Console.WriteLine(tag);
            string ext = Path.GetExtension(tag);
            if(ext == ".lnk")
            {
                IWshRuntimeLibrary.WshShell shell = new IWshRuntimeLibrary.WshShell();
                IWshRuntimeLibrary.IWshShortcut shortcut = (IWshRuntimeLibrary.IWshShortcut)shell.CreateShortcut(tag);
                string targetPath = shortcut.TargetPath.ToString();
                try
                {
                    System.Diagnostics.Process.Start(targetPath);
                }
                catch (Exception ex)
                {
                    try
                    {
                        System.Diagnostics.Process.Start(targetPath.Replace("C:\\Program Files (x86)\\", "C:\\Program Files\\"));
                    }//x86を消して実行
                    catch (Exception eex)
                    {
                        try
                        {
                            System.Diagnostics.Process.Start(tag);
                        }
                        catch (Exception exe)
                        {
                           
                        }

                    }
                }
            }
            else
            {
                try
                {
                    System.Diagnostics.Process.Start(tag);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }//通常の実行
        }
        private void OnPressOpen(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("EXPLORER.EXE", $@"shortcuts");
            Console.WriteLine($"{Directory.GetCurrentDirectory() + "\\shortcuts\\"}");
        }
        private void OnPressReload(object sender,EventArgs e)
        {
            Application.Restart();
        }


    }
}
