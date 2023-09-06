using AudioSwitch.COM;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AudioSwitch
{
    public partial class Form1 : Form
    {
        HotkeyManager hotkeyManager;
        public Form1()
        {
            InitializeComponent();
            this.hotkeyManager = new HotkeyManager(this.Handle);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.hotkeyManager.Add(0, NativeAPI.MOD_CONTROL | NativeAPI.MOD_ALT, Keys.F11, this.NextAudioDevice);
            if (!this.hotkeyManager.RegisterHotkeys())
            {
                MessageBox.Show("Hotkey register failed", "AudioSwitch Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
        }
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == NativeAPI.WM_HOTKEY)
                this.hotkeyManager.Dispatch((uint)m.WParam);
            else
                base.WndProc(ref m);
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            this.notifyIcon1.Icon = this.Icon;
            this.notifyIcon1.Visible = true;
            this.Hide();
        }

        private void NextAudioDevice()
        {
            using(var deviceEnumerator = new COMHolder<IMMDeviceEnumerator>(new MMDeviceEnumerator() as IMMDeviceEnumerator))
            {
                string activeId;
                using (var defaultDevice = new COMHolder<IMMDevice>(deviceEnumerator.Value.GetDefaultAudioEndpoint(EDataFlow.Render, ERole.Multimedia)))
                {
                    activeId = defaultDevice.Value.GetId();
                }

                string firstId = null, prevId = null, selectedId = null;
                using (var endpoints = new COMHolder<IMMDeviceCollection>(deviceEnumerator.Value.EnumAudioEndpoints(EDataFlow.Render, 1)))
                {
                    for (int i = 0; i < endpoints.Value.GetCount(); i++)
                    {
                        using (var item = new COMHolder<IMMDevice>(endpoints.Value.Item(i)))
                        {
                            string id = item.Value.GetId();

                            if (string.IsNullOrEmpty(firstId))
                            {
                                firstId = id;
                            }

                            if (prevId == activeId)
                            {
                                selectedId = id;
                                break;
                            }

                            prevId = id;
                        }
                    }
                }

                if (string.IsNullOrEmpty(selectedId))
                {
                    selectedId = firstId;
                }

                if (string.IsNullOrEmpty(selectedId))
                {
                    return;
                }

                new CPolicyConfigVistaClient().SetDefaultDevice(selectedId);
            }
        }

        private void notifyIcon1_MouseClick(object sender, MouseEventArgs e)
        {
            NextAudioDevice();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            notifyIcon1.Visible = false;
            Application.Exit();
        }

        private void audioSwitchByInndyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process.Start("https://github.com/Inndy/AudioSwitch");
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.hotkeyManager.UnregisterHotkeys();
        }
    }
}
