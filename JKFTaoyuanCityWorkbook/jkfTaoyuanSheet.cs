// author: alvin.lin@outlook.com
// datetime: 2024-2-18 21:03pm
// license: all rights are reversed.

using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Globalization;


namespace JKFTaoyuanCityWorkbook
{
    public partial class jkfTaoyuanSheet
    {
        private void 工作表1_Startup(object sender, System.EventArgs e)
        {
            FirefoxOptions options = new FirefoxOptions();
            options.AddArgument("--headless");
            options.AddArgument("--disable-gpu");
            options.AddArgument("--no-sandbox");
            string taoyuan_area = "<url>";
            Uri uri = new Uri(taoyuan_area);
            string scheme = uri.Scheme;
            string netloc = uri.Host;
            string site_name = $"{scheme}://{netloc}/";
            string[] headers = new string[] { "筆", "頭像", "頭像連結", "置頂", "地區", "地區連結", "現在有空", "分區", "標題", "標題連結", "回覆", "觀看", "發文者", "發文者連結", "發文日期", "最後回覆", "回覆者連結", "回覆文連結", "回覆日期時間" };
            string[] area_zones = new string[] { "南門", "復興路", "大興西路", "車站", "三民", "藝文", "桃園", "八德", "內壢", "鶯歌", "中壢", "蘆竹", "南崁", "大園", "楊梅", "觀音", "龜山", "長庚", "林口", "三峽", "龍潭" };
            // 總共要抓幾頁
            int capture_pages = 1;
            // 目前在第幾行
            int currentRow = 1;
            // 寫入標題
            for (int i = 0; i < headers.Length; i++)
            {
                this.Cells[currentRow, i + 1].Value = headers[i];
            }
            // 寫入到標題的每一欄
            int headersLength = headers.Length;
            char endColumnLetter = (char)('A' + headersLength - 1);
            Excel.Range titleRange = this.Range[$"A1:{endColumnLetter}1"];
            // 設定背景顏色
            titleRange.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(128, 96, 0));
            // 設定字體顏色
            titleRange.Font.Color = ColorTranslator.ToOle(Color.White);
            // 設定字體加粗
            titleRange.Font.Bold = true;
            // 在 titleRange 範圍啟用自動篩選
            titleRange.AutoFilter(1);
            currentRow += 1;
            // options
            using (IWebDriver browser = new FirefoxDriver(options))
            {
                browser.Navigate().GoToUrl("https://www.jkforum.net/p/type-1128-1947.html");
                // 等待並處理彈窗...
                WebDriverWait wait = new WebDriverWait(browser, TimeSpan.FromSeconds(10));
                try
                {
                    var webElements = wait.Until(driver => driver.FindElement(By.Id("fd_page_bottom")));
                }
                catch
                {
                    System.Diagnostics.Debug.WriteLine($"第 {currentRow} 列等待並處理彈窗錯誤");
                }
                // 處理彈窗
                IWebElement dontRemind = browser.FindElement(By.XPath("//*[@id='periodaggre18']"));
                dontRemind.Click();
                // 處理彈窗
                IWebElement yesOver18 = browser.FindElement(By.XPath("//*[@id='fwin_dialog_submit']"));
                yesOver18.Click();
                try
                {
                    var webElements = wait.Until(driver => driver.FindElement(By.ClassName("nxt")));
                }
                catch
                {
                    System.Diagnostics.Debug.WriteLine($"第 {currentRow} 列處理彈窗錯誤");
                }
                // 資料抓取和處理...
                for (int nextPage = 0; nextPage < capture_pages; nextPage++)
                {
                    var table = wait.Until(driver => driver.FindElement(By.Id("threadlisttableid")));
                    var trs = table.FindElements(By.TagName("tr"));
                    if (trs.Count > 0) // Ensure it's a data row
                    {
                        // 分別從 HTML 表格裡的每一列 row 裡旳 tr 中擷取出 th 和 td 的內容
                        foreach (var tr in trs)
                        {
                            // 筆
                            this.Cells[currentRow, Array.IndexOf(headers, "筆") + 1] = currentRow;
                            // 變更雙數列顏色
                            if (currentRow % 2 == 1)
                            {
                                Excel.Range range = (Excel.Range)this.Rows[currentRow];
                                range.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(255, 242, 204));
                            }
                            // 從 tr 中擷取出 th 的內容
                            var ths = tr.FindElements(By.TagName("th"));
                            if (ths.Count == 1)
                            {
                                // 確認是否 tr 的父節點為分格線
                                var trParentId = tr.FindElement(By.XPath("..")).GetAttribute("id");
                                if (trParentId == "separatorline_top")
                                {
                                    continue;
                                }
                                // 從 th 中擷取出第一個和唯一個內容
                                foreach (var th in ths)
                                {
                                    var th_span_a = th.FindElement(By.TagName("a"));
                                    if (th_span_a != null)
                                    {
                                        // 頭像
                                        try
                                        {
                                            var thumbnail_image = th_span_a.FindElement(By.TagName("img"));
                                            if (thumbnail_image != null)
                                            {
                                                // 讀取並處理單元格值
                                                var columnIndex = Array.IndexOf(headers, "頭像") + 1;
                                                string cellStrValue = thumbnail_image.GetAttribute("src");
                                                Excel.Range cell = this.Cells[currentRow, columnIndex] as Excel.Range;
                                                if (!string.IsNullOrEmpty(cellStrValue))
                                                {
                                                    cell.Value = cellStrValue;
                                                    // 新增超連結
                                                    this.Hyperlinks.Add(cell, cellStrValue, "", cellStrValue);
                                                }
                                                else
                                                {
                                                    // 處理錯誤值
                                                    System.Diagnostics.Debug.WriteLine($"第 {currentRow} 列的頭像錯誤");
                                                }
                                            }
                                            // 頭像連結
                                            var thumbnail_href = th_span_a.GetAttribute("href");
                                            if (thumbnail_href != null)
                                            {
                                                var columnIndex = Array.IndexOf(headers, "頭像連結") + 1;
                                                Excel.Range cell = this.Cells[currentRow, columnIndex] as Excel.Range;
                                                if (!string.IsNullOrEmpty(thumbnail_href))
                                                {
                                                    cell.Value = thumbnail_href;
                                                    // 新增超連結
                                                    this.Hyperlinks.Add(cell, thumbnail_href, "", thumbnail_href);
                                                }
                                                else
                                                {
                                                    // 處理錯誤值
                                                    System.Diagnostics.Debug.WriteLine($"第 {currentRow} 列的頭像連結錯誤");
                                                }
                                            }
                                        }
                                        catch
                                        {
                                            System.Diagnostics.Debug.WriteLine($"第 {currentRow} 列頭像錯誤");
                                        }
                                    }
                                    var divs = th.FindElements(By.TagName("div"));
                                    if (divs.Count == 3)
                                    {
                                        for (int i = 0; i < divs.Count; i++)
                                        {
                                            switch (i)
                                            {
                                                case 0:
                                                    // 置頂
                                                    try
                                                    {
                                                        var pin_to_top = divs[i].FindElement(By.TagName("img")).GetAttribute("src");
                                                        if (!string.IsNullOrEmpty(pin_to_top))
                                                        {
                                                            this.Cells[currentRow, Array.IndexOf(headers, "置頂") + 1] = site_name + pin_to_top;
                                                        }
                                                    }
                                                    catch
                                                    {
                                                        System.Diagnostics.Debug.WriteLine($"第 {currentRow} 列置頂錯誤");
                                                    }
                                                    // 地區
                                                    try
                                                    {
                                                        var area_name = divs[i].FindElement(By.TagName("a")).Text;
                                                        if (!string.IsNullOrEmpty(area_name))
                                                        {
                                                            this.Cells[currentRow, Array.IndexOf(headers, "地區") + 1] = area_name;
                                                        }
                                                    }
                                                    catch
                                                    {
                                                        System.Diagnostics.Debug.WriteLine($"第 {currentRow} 列地區錯誤");
                                                    }
                                                    // 地區連結
                                                    try
                                                    {
                                                        var area_href = divs[i].FindElement(By.TagName("a")).GetAttribute("href");
                                                        if (!string.IsNullOrEmpty(area_href))
                                                        {
                                                            this.Cells[currentRow, Array.IndexOf(headers, "地區連結") + 1] = area_href;
                                                        }
                                                    }
                                                    catch
                                                    {
                                                        System.Diagnostics.Debug.WriteLine($"第 {currentRow} 列地區連結錯誤");
                                                    }
                                                    // 現在有空
                                                    try
                                                    {
                                                        var div_images = divs[i].FindElements(By.TagName("img"));
                                                        if (div_images.Count == 2)
                                                        {
                                                            var availibility = site_name + div_images[1].GetAttribute("src");
                                                            this.Cells[currentRow, Array.IndexOf(headers, "現在有空") + 1] = availibility;
                                                        }
                                                    }
                                                    catch
                                                    {
                                                        System.Diagnostics.Debug.WriteLine($"第 {currentRow} 列地區連結錯誤");
                                                    }
                                                    // 標題、標題連結和分區
                                                    try
                                                    {
                                                        var subjectColumnIndex = Array.IndexOf(headers, "標題") + 1;
                                                        var subjHrefColumnIndex = Array.IndexOf(headers, "標題連結") + 1;
                                                        var zoneColumnIndex = Array.IndexOf(headers, "分區") + 1;
                                                        var div_a_sxst = divs[i].FindElements(By.CssSelector("a.s.xst"));
                                                        if (div_a_sxst.Count == 1 && div_a_sxst[0].Text != "")
                                                        {
                                                            // 標題
                                                            var post_title = div_a_sxst[0].Text;
                                                            // 標題連結
                                                            var post_href = div_a_sxst[0].GetAttribute("href");
                                                            if (!string.IsNullOrEmpty(post_title)) // 標題
                                                            {
                                                                // 移除開頭的 '+' 或 '='
                                                                if (post_title.StartsWith("+"))
                                                                {
                                                                    post_title = post_title.Substring(1);
                                                                }
                                                                if (post_title.StartsWith("="))
                                                                {
                                                                    post_title = post_title.Substring(1);
                                                                }
                                                                if (!string.IsNullOrEmpty(post_href)) // 標題連結
                                                                {
                                                                    Excel.Range cell = this.Cells[currentRow, subjectColumnIndex] as Excel.Range; // 標題
                                                                    cell.Value = post_title;
                                                                    this.Cells[currentRow, subjHrefColumnIndex] = post_href; // 標題連結
                                                                    // 新增標題連結到標題的超連結
                                                                    this.Hyperlinks.Add(cell, post_href, "", post_title);
                                                                }
                                                                // 分區
                                                                string area_zone = area_zones.FirstOrDefault(zone => post_title.Contains(zone));
                                                                if (!string.IsNullOrEmpty(area_zone))
                                                                {
                                                                    this.Cells[currentRow, zoneColumnIndex] = area_zone;
                                                                }
                                                            }
                                                        }
                                                    }
                                                    catch
                                                    {
                                                        System.Diagnostics.Debug.WriteLine($"第 {currentRow} 列標題、標題連結和分區錯誤");
                                                    }
                                                    break;
                                                case 1:
                                                    // 這部份不需要處理直接跳過
                                                    break;
                                                case 2:
                                                    // 處理 回覆(posted) 和 觀看(viewed) 的值
                                                    var div_text = divs[i].Text;
                                                    if (!string.IsNullOrEmpty(div_text))
                                                    {
                                                        // 移除空白字元
                                                        div_text = div_text.Replace(" ", "").Replace("\t", "").Replace("\n", "");
                                                        div_text = div_text.TrimEnd('/');

                                                        // 判斷是否只包含一個 '/'
                                                        if (div_text.Count(f => f == '/') == 1)
                                                        {
                                                            // 在這裡處理 posted 和 viewed 的值
                                                            var parts = div_text.Split('/');
                                                            // 回覆
                                                            var posted = parts[0];
                                                            this.Cells[currentRow, Array.IndexOf(headers, "回覆") + 1] = posted;
                                                            // 觀看
                                                            var viewed = parts[1];
                                                            this.Cells[currentRow, Array.IndexOf(headers, "觀看") + 1] = viewed;
                                                        }
                                                    }
                                                    break;
                                            }
                                        }
                                    }
                                }
                            }
                            // 從 tr 中擷取出 td 的內容
                            var tds = tr.FindElements(By.TagName("td"));
                            if (tds.Count == 3)
                            {
                                // 發文者、發文者連結、發文日期、最後回覆、回覆者連結、回覆文連結 和 回覆日期時間
                                for (int i = 0; i < tds.Count; i++)
                                {
                                    switch (i)
                                    {
                                        case 0:
                                            // 發文者、發文者連結和發文日期
                                            var td_cite_a = tds[i].FindElement(By.TagName("cite")).FindElement(By.TagName("a"));
                                            if (td_cite_a != null)
                                            {
                                                // 發文者
                                                var authorColumnIndex = Array.IndexOf(headers, "發文者") + 1;
                                                var author_name = td_cite_a.Text;
                                                // 發文者連結
                                                var authorHrefColumnIndex = Array.IndexOf(headers, "發文者連結") + 1;
                                                var auther_href = td_cite_a.GetAttribute("href");
                                                if (!string.IsNullOrEmpty(author_name) && !string.IsNullOrEmpty(auther_href)) // 發文者 & 發文者連結
                                                {
                                                    Excel.Range cell = this.Cells[currentRow, authorColumnIndex] as Excel.Range; // 發文者
                                                    cell.Value = author_name;
                                                    this.Cells[currentRow, authorHrefColumnIndex] = auther_href; // 發文者連結
                                                    // 新增發文者連結到發文者的超連結
                                                    this.Hyperlinks.Add(cell, auther_href, "", author_name);
                                                }
                                            }
                                            var td_em_span = tds[i].FindElement(By.TagName("em")).FindElement(By.TagName("span"));
                                            if (td_em_span != null)
                                            {
                                                // 發文日期
                                                var auther_date = td_em_span.Text;
                                                this.Cells[currentRow, Array.IndexOf(headers, "發文日期") + 1] = auther_date;
                                            }
                                            break;
                                        case 1:
                                            // 這部份不需要處理直接跳過
                                            break;
                                        case 2:
                                            td_cite_a = tds[i].FindElement(By.TagName("cite")).FindElement(By.TagName("a"));
                                            if (td_cite_a != null)
                                            {
                                                // 最後回覆
                                                var poster_name = td_cite_a.Text;
                                                this.Cells[currentRow, Array.IndexOf(headers, "最後回覆") + 1] = poster_name;
                                                // 回覆者連結
                                                var poster_href = site_name + td_cite_a.GetAttribute("href");
                                                this.Cells[currentRow, Array.IndexOf(headers, "回覆者連結") + 1] = poster_href;
                                            }
                                            var td_em_a = tds[i].FindElement(By.TagName("em")).FindElement(By.TagName("a"));
                                            if (td_em_a != null)
                                            {
                                                // 回覆文連結
                                                var poster_url = site_name + td_em_a.GetAttribute("href");
                                                this.Cells[currentRow, Array.IndexOf(headers, "回覆文連結") + 1] = poster_url;
                                                // 回覆日期時間
                                                var poster_datetime = td_em_a.Text.Replace("\xa0", "");
                                                this.Cells[currentRow, Array.IndexOf(headers, "回覆日期時間") + 1] = poster_datetime;
                                            }
                                            break;
                                    }
                                }
                            }
                            currentRow += 1;
                            System.Diagnostics.Debug.WriteLine($"第 {currentRow} 列處理完成。");
                        }


                        // Handling pagination
                        try
                        {
                            var nextPageButton = browser.FindElement(By.ClassName("nxt"));
                            nextPageButton.Click();
                            wait.Until(driver => driver.FindElement(By.Id("fd_page_bottom")));
                        }
                        catch (NoSuchElementException)
                        {
                            System.Diagnostics.Debug.WriteLine("No more pages.");
                        }
                        catch (TimeoutException)
                        {
                            System.Diagnostics.Debug.WriteLine("Timeout waiting for next page.");
                        }
                    }
                }
                browser.Quit();
                this.UsedRange.Font.Name = "Yu Gothic"; // 設定字型
                this.UsedRange.Font.Size = 10; // 設定字體大小
                this.UsedRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter; // 設定垂直對齊方式為居中
                // 獲取當前使用者的桌面路徑
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                // 設定每一欄的寬度
                int[] widthAlignment = new int[] { 4, 8, 8, 4, 4, 7, 7, 5, 72, 8, 5, 14, 10, 11, 13, 8, 15, 15, 18 };
                for (int i = 0; i < widthAlignment.Length; i++)
                {
                    this.Columns[i + 1].ColumnWidth = widthAlignment[i];
                }
                string filePath = Path.Combine(desktopPath, "jkf_taoyuan.xlsx");
                this.Application.ActiveWorkbook.SaveAs(filePath, Excel.XlFileFormat.xlOpenXMLWorkbook);
            }
        }

        private void 工作表1_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(工作表1_Startup);
            this.Shutdown += new System.EventHandler(工作表1_Shutdown);
        }

        #endregion

    }
}
