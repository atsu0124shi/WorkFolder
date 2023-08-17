using System;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using WebDriverManager;
using WebDriverManager.DriverConfigs.Impl;
using WebDriverManager.Helpers;
using SeleniumExtras.WaitHelpers;
using System.Runtime.CompilerServices;
using AngleSharp.Dom;
using System.Threading.Tasks;
using System.IO;
using System.Linq;

namespace Utility
{
    public class File
    {
        /// <summary>
        /// 非同期で新規タブを開いてクロムのヘルプサイトを開いて自動アップデートする。
        /// </summary>
        /// <param name="driver"></param>
        public static async Task Update (IWebDriver driver)
        {
            driver.Navigate().GoToUrl(@"chrome://settings/help");
        }

        /// <summary>
        /// 非同期処理で指定したディレクトリ（サブ含む）を削除する。
        /// </summary>
        public static async Task DirDeleteAsync (string dir)
        {
            await Task.Run(() =>
            {
                var directoryInfo = new DirectoryInfo(dir).GetDirectories();

                if (directoryInfo.Count() > 2)
                {
                    var delDir = directoryInfo.OrderBy(a => a.LastWriteTime).FirstOrDefault();

                    delDir.Delete(true);
                }
            });

        }

    }
}