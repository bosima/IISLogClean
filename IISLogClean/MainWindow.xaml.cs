using NLog;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;

namespace IISLogClean
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();
        private bool isRunning = false;
        private DispatcherTimer planTimer = null;
        private DateTime? lastDayRunTime = null;
        private DateTime? lastWeekRunTime = null;
        private DateTime? lastMonthRunTime = null;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            if (isRunning)
            {
                MessageBox.Show("正在处理，请稍后！");
            }

            string logBasePath = this.txtLogPath.Text;
            string strLogExpireDay = this.txtLogExpireDay.Text;
            int logExpireDay = 0;

            if (string.IsNullOrWhiteSpace(logBasePath))
            {
                MessageBox.Show("【日志文件夹】必填！");
                return;
            }

            if (string.IsNullOrWhiteSpace(strLogExpireDay))
            {
                MessageBox.Show("【几天前日志】必填！");
                return;
            }

            if (!int.TryParse(strLogExpireDay, out logExpireDay))
            {
                MessageBox.Show("【几天前日志】请填写数字！");
                return;
            }

            if (!Directory.Exists(logBasePath))
            {
                MessageBox.Show("【日志文件夹】不存在！");
                return;
            }

            DoClean(logBasePath, logExpireDay);
        }

        private void DoClean(string logBasePath, int logExpireDay)
        {
            isRunning = true;

            Task.Factory.StartNew(() =>
            {
                txtProgress.Dispatcher.Invoke(new Action(() =>
                {
                    txtProgress.Clear();
                }));

                int successCount = 0;
                int failCount = 0;
                string[] subDirectories = null;

                try
                {
                    subDirectories = Directory.GetDirectories(logBasePath);
                }
                catch (Exception ex)
                {
                    RenderProgress("目录访问失败：" + ex.Message);
                }

                if (subDirectories != null && subDirectories.Length > 0)
                {
                    foreach (var dirName in subDirectories)
                    {
                        string siteLogPath = System.IO.Path.Combine(logBasePath, dirName);

                        // 获取子目录下的文件
                        string[] siteLogFiles;
                        try
                        {
                            siteLogFiles = Directory.GetFiles(siteLogPath);
                        }
                        catch (Exception ex)
                        {
                            RenderProgress("获取目录下的文件失败：" + ex.Message);
                            continue;
                        }

                        if (siteLogFiles.Length > 0)
                        {
                            foreach (var logName in siteLogFiles)
                            {
                                var logPath = System.IO.Path.Combine(siteLogPath, logName);
                                var logLastWriteTime = File.GetLastWriteTime(logPath);
                                if (logLastWriteTime.AddDays(logExpireDay) < DateTime.Now)
                                {
                                    try
                                    {
                                        File.Delete(logPath);
                                        successCount++;
                                    }
                                    catch (Exception ex)
                                    {
                                        RenderProgress("删除文件失败：" + ex.Message);
                                        failCount++;
                                    }
                                }
                            }
                        }

                        RenderProgress(dirName + "处理完毕。");
                    }
                }

                RenderProgress(string.Format("全部处理完毕，成功{0}条，失败{1}条。", successCount, failCount));
                isRunning = false;
            });
        }

        private void RenderProgress(string content)
        {
            logger.Info(content);

            txtProgress.Dispatcher.Invoke(new Action(() =>
            {
                txtProgress.AppendText(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") + content + Environment.NewLine);
            }));
        }

        private void comboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbPeriodItem != null)
            {
                cmbPeriodItem.Items.Clear();

                var type = cmbPeriodType.SelectedIndex;
                if (type == 0)
                {
                    cmbPeriodItem.IsEnabled = false;
                }
                else if (type == 1)
                {
                    cmbPeriodItem.IsEnabled = true;
                    SetPeriodItemWeek();
                }
                else if (type == 2)
                {
                    cmbPeriodItem.IsEnabled = true;
                    SetPeriodItemMonth();
                }
            }
        }

        private void SetPeriodItemMonth()
        {
            for (var i = 1; i <= 29; i++)
            {
                cmbPeriodItem.Items.Add(i + "日");
            }

            cmbPeriodItem.SelectedIndex = 0;
        }

        private void SetPeriodItemWeek()
        {
            var item1 = new ComboBoxItem();
            item1.Content = "星期一";
            item1.DataContext = 1;
            cmbPeriodItem.Items.Add(item1);

            var item2 = new ComboBoxItem();
            item2.Content = "星期二";
            item2.DataContext = 2;
            cmbPeriodItem.Items.Add(item2);

            var item3 = new ComboBoxItem();
            item3.Content = "星期三";
            item3.DataContext = 3;
            cmbPeriodItem.Items.Add(item3);

            var item4 = new ComboBoxItem();
            item4.Content = "星期四";
            item4.DataContext = 4;
            cmbPeriodItem.Items.Add(item4);

            var item5 = new ComboBoxItem();
            item5.Content = "星期五";
            item5.DataContext = 5;
            cmbPeriodItem.Items.Add(item5);

            var item6 = new ComboBoxItem();
            item6.Content = "星期六";
            item6.DataContext = 6;
            cmbPeriodItem.Items.Add(item6);

            var item7 = new ComboBoxItem();
            item7.Content = "星期日";
            item7.DataContext = 0;
            cmbPeriodItem.Items.Add(item7);

            cmbPeriodItem.SelectedIndex = 0;
        }

        private void SetPeriodTime()
        {
            for (var i = 0; i <= 23; i++)
            {
                cmbPeriodTimeHour.Items.Add(i.ToString().PadLeft(2, '0') + "点");
            }

            cmbPeriodTimeHour.SelectedIndex = 3;

            for (var j = 0; j <= 55; j = j + 5)
            {
                cmbPeriodTimeMinute.Items.Add(j.ToString().PadLeft(2, '0') + "分");
            }

            cmbPeriodTimeMinute.SelectedIndex = 2;
        }

        private void button2_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new System.Windows.Forms.FolderBrowserDialog { SelectedPath = txtLogPath.Text };
            var res = dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK;
            if (!res) return;
            txtLogPath.Text = dlg.SelectedPath;
        }

        private void IISLogCleanMainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            SetPeriodTime();
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            string logBasePath = this.txtLogPath.Text;
            string strLogExpireDay = this.txtLogExpireDay.Text;
            int logExpireDay = 0;

            if (string.IsNullOrWhiteSpace(logBasePath))
            {
                MessageBox.Show("【日志文件夹】必填！");
                return;
            }

            if (string.IsNullOrWhiteSpace(strLogExpireDay))
            {
                MessageBox.Show("【几天前日志】必填！");
                return;
            }

            if (!int.TryParse(strLogExpireDay, out logExpireDay))
            {
                MessageBox.Show("【几天前日志】请填写数字！");
                return;
            }

            if (!Directory.Exists(logBasePath))
            {
                MessageBox.Show("【日志文件夹】不存在！");
                return;
            }

            txtPlanStatus.Content = "清理计划已于" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "启动";
            btnStartPlan.IsEnabled = false;

            planTimer = new DispatcherTimer();
            planTimer.Interval = new TimeSpan(0, 0, 1);
            planTimer.Tick += PlanTimer_Tick;
            planTimer.Start();
        }

        private void PlanTimer_Tick(object sender, EventArgs e)
        {
            var now = DateTime.Now;
            var type = cmbPeriodType.SelectedIndex;
            if (type == 0)
            {
                var periodTimeHour = cmbPeriodTimeHour.SelectedValue.ToString();
                var periodTimeMinute = cmbPeriodTimeMinute.SelectedValue.ToString();
                var periodTime = new DateTime(now.Year, now.Month, now.Day, int.Parse(periodTimeHour.Replace("点", "")), int.Parse(periodTimeMinute.Replace("分", "")), 0);

                if (now.ToString("yyyy-MM-dd HH:mm") == periodTime.ToString("yyyy-MM-dd HH:mm"))
                {
                    // 如果上次运行的时间不是今天，则说明今天还没有运行过
                    if (lastDayRunTime.HasValue && lastDayRunTime.Value.ToString("yyyy-MM-dd") != now.ToString("yyyy-MM-dd"))
                    {
                        lastDayRunTime = null;
                    }

                    if (lastDayRunTime == null)
                    {
                        while (isRunning)
                        {
                            Thread.Sleep(1000);
                        }

                        lastDayRunTime = DateTime.Now;

                        string logBasePath = this.txtLogPath.Text;
                        string strLogExpireDay = this.txtLogExpireDay.Text;
                        int logExpireDay = int.Parse(strLogExpireDay);

                        DoClean(logBasePath, logExpireDay);
                    }
                }
            }
            else if (type == 1)
            {
                var periodWeekIndex = (int)((ComboBoxItem)cmbPeriodItem.SelectedItem).DataContext;
                var weekIndex = Convert.ToInt32(DateTime.Now.DayOfWeek.ToString("d"));
                var periodTimeHour = cmbPeriodTimeHour.SelectedValue.ToString().Replace("点", "");
                var periodTimeMinute = cmbPeriodTimeMinute.SelectedValue.ToString().Replace("分", "");

                if (periodWeekIndex == weekIndex && now.ToString("HH:mm") == periodTimeHour + ":" + periodTimeMinute)
                {
                    // 如果上次运行的时间不是本周，则说明本周还没有运行过
                    if (lastWeekRunTime.HasValue && GetWeekOfYear(lastWeekRunTime.Value) != GetWeekOfYear(now))
                    {
                        lastWeekRunTime = null;
                    }

                    if (lastWeekRunTime == null)
                    {
                        while (isRunning)
                        {
                            Thread.Sleep(1000);
                        }

                        lastWeekRunTime = DateTime.Now;

                        string logBasePath = this.txtLogPath.Text;
                        string strLogExpireDay = this.txtLogExpireDay.Text;
                        int logExpireDay = int.Parse(strLogExpireDay);

                        DoClean(logBasePath, logExpireDay);
                    }
                }
            }
            else if (type == 2)
            {
                var periodDayIndex = int.Parse(cmbPeriodItem.SelectedItem.ToString().Replace("日", "").Replace("0", ""));
                var dayIndex = DateTime.Now.Day;
                var periodTimeHour = cmbPeriodTimeHour.SelectedValue.ToString().Replace("点", "");
                var periodTimeMinute = cmbPeriodTimeMinute.SelectedValue.ToString().Replace("分", "");

                if (periodDayIndex == dayIndex && now.ToString("HH:mm") == periodTimeHour + ":" + periodTimeMinute)
                {
                    // 如果上次运行的时间不是本月，则说明本月还没有运行过
                    if (lastMonthRunTime.HasValue && lastMonthRunTime.Value.Month != now.Month)
                    {
                        lastMonthRunTime = null;
                    }

                    if (lastMonthRunTime == null)
                    {
                        while (isRunning)
                        {
                            Thread.Sleep(1000);
                        }

                        lastMonthRunTime = DateTime.Now;

                        string logBasePath = this.txtLogPath.Text;
                        string strLogExpireDay = this.txtLogExpireDay.Text;
                        int logExpireDay = int.Parse(strLogExpireDay);

                        DoClean(logBasePath, logExpireDay);
                    }
                }
            }
        }

        /// <summary>
        /// 获取指定日期，在为一年中为第几周
        /// </summary>
        /// <param name="dt">指定时间</param>
        /// <reutrn>返回第几周</reutrn>
        private static int GetWeekOfYear(DateTime dt)
        {
            GregorianCalendar gc = new GregorianCalendar();
            int weekOfYear = gc.GetWeekOfYear(dt, CalendarWeekRule.FirstDay, DayOfWeek.Sunday);
            return weekOfYear;
        }
    }
}
