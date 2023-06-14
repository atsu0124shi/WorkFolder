using System;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using WebDriverManager;
using WebDriverManager.DriverConfigs.Impl;
using WebDriverManager.Helpers;
using SeleniumExtras.WaitHelpers;
using System.Threading.Tasks;
using System.Threading;
using Utility;
using System.IO;
using ExcelDataReader;
using System.Text;
using System.Data;
using static System.Net.Mime.MediaTypeNames;
using System.Runtime.CompilerServices;
using System.Linq;

namespace Luupy自動貼り換え
{
    class Luupy
    {
        static void Main(string[] args)
        {
            // 定数（ログイン情報）
            const string mail = "email";
            const string password = "password";
            const string logIn = "Login";
            const string kokokuKanri = "広告管理";
            const string kokokuwaku2 = "ad02";
            const string updateButton = "更新する";
            const string logOut = "ログアウト";
            const string filePath = @"..\LUPPY管理簿.xlsx";


            Console.WriteLine("ChromeDriverを取得中");

            // インストールされているバージョン
            new DriverManager().SetUpDriver(new ChromeConfig(), VersionResolveStrategy.MatchingBrowser);

            Console.WriteLine("ChromeDriverを取得終了");

            // Webドライバーのインスタンス化
            IWebDriver driver = new ChromeDriver();

            WebDriverWait wait = new WebDriverWait(driver, new TimeSpan(0, 0, 5));

            // 非同期でクロムのヘルプを開いて自動アップデートする
            Task.Run(async () =>
            {
                await Utility.File.Update(driver);
            });

            // 会員サイトから今月の画像URL(広告枠2)を取得して、LUUPY管理簿にセット
            string settingValue = Utility.LuupyExcel.Update(driver, filePath);

            // Data関連を定義
            DataSet dataset = new DataSet();
            DataTable luppyDataTable = new DataTable();

            #region Luupyエクセルから情報を取得
            using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var streamReader = new StreamReader(stream))
            {
                IExcelDataReader reader;
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

                reader = ExcelReaderFactory.CreateReader(streamReader.BaseStream, new ExcelReaderConfiguration()
                {
                    FallbackEncoding = Encoding.GetEncoding("Shift_JIS")
                });

                dataset = reader.AsDataSet();

                // DataTableに設定
                luppyDataTable = dataset.Tables["一覧"];

                reader.Dispose();
            }
            #endregion

            // 広告枠2の設定値を設定
            //string settingValue = settingDataTable.Rows[1][2].ToString();

            // ログインして広告設定値を変更する
            // ループする処理はここから
            for (int i = 0; i < luppyDataTable.Rows.Count; i++)
            {
                // 最初の行は見出し行なので飛ばす
                if (i == 0)
                {
                    continue;
                }

                // 新規タブを開く
                driver.SwitchTo().NewWindow(WindowType.Tab);

                // ログイン情報を設定
                string url = luppyDataTable.Rows[i][3].ToString();
                string mailAddress = luppyDataTable.Rows[i][4].ToString();
                string pass = luppyDataTable.Rows[i][5].ToString();

                // URLに移動します。
                driver.Navigate().GoToUrl(url);

                // ログイン情報設定
                wait.Until(ExpectedConditions.ElementExists(By.Id(mail)));
                driver.FindElement(By.Id(mail)).SendKeys(mailAddress);
                driver.FindElement(By.Id(password)).SendKeys(pass);
                driver.FindElement(By.XPath($"//*[contains(text(), '{logIn}')]")).Click();

                // 広告管理に移動
                wait.Until(ExpectedConditions.ElementExists(By.XPath($"//*[contains(text(), '{kokokuKanri}')]")));
                driver.FindElement(By.XPath($"//*[contains(text(), '{kokokuKanri}')]")).Click();

                // 広告枠2に設定
                wait.Until(ExpectedConditions.ElementExists(By.Name(kokokuwaku2)));
                driver.FindElement(By.Name(kokokuwaku2)).Clear();
                driver.FindElement(By.Name(kokokuwaku2)).SendKeys(settingValue);

                // 更新ボタンを押下
                wait.Until(ExpectedConditions.ElementExists(By.XPath($"//*[contains(text(), '{updateButton}')]")));
                driver.FindElement(By.XPath($"//*[contains(text(), '{updateButton}')]")).Click();

                // ログアウト処理
                wait.Until(ExpectedConditions.ElementExists(By.XPath($"//*[contains(text(), '{logOut}')]")));
                driver.FindElement(By.XPath($"//*[contains(text(), '{logOut}')]")).Click();

                // タブを閉じる
                driver.Close();

                // 最初のタブを選択
                driver.SwitchTo().Window(driver.WindowHandles.First());
            }

            // 処理終了後、閉じる
            driver.Quit();
        }
    }
}