using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using System.Text;
using CommandLine;
using System.Text.RegularExpressions;
using System.IO;
using System.Net.Http;
using System.Web;
using Newtonsoft.Json;
using notify_slack_of_web_meeting.cli.Settings;
using notify_slack_of_web_meeting.cli.SlackChannels;
using notify_slack_of_web_meeting.cli.WebMeetings;
using JsonSerializer = System.Text.Json.JsonSerializer;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace notify_slack_of_web_meeting.cli
{
    class Program
    {
        /// <summary>
        /// HTTPクライアント
        /// </summary>
        private static HttpClient s_HttpClient = new HttpClient();

        [Verb("setting", HelpText = "Register Slack channel information and create a configuration file.")]
        public class SettingOptions
        {
            [Option('u', "url", HelpText = "The web service endpoint url.", Required = true)]
            public string EndpointUrl { get; set; }

            [Option('n', "name", HelpText = "The Slack channel name.", Required = true)]
            public string Name { get; set; }

            [Option('w', "webhookurl", HelpText = "The web hook url. (Slack incoming webhook)", Required = true)]
            public string WebhookUrl { get; set; }

            [Option('r', "register", HelpText = "The registered name.", Required = true)]
            public string RegisteredBy { get; set; }

            [Option('f', "filepath", HelpText = "Ourput setting file path.", Default = "./setting.json")]
            public string Filepath { get; set; }

        }
        [Verb("register", HelpText = "Register the web conference information to be notified.")]
        public class RegisterOptions
        {
            [Option('f', "filepath", HelpText = "Input setting file path.", Default = "./setting.json")]
            public string Filepath { get; set; }
        }

        static int Main(string[] args)
        {
            Func<SettingOptions, int> RunSettingAndReturnExitCode = opts =>
            {
                Console.WriteLine("Run Setting");

                #region 引数の値でSlackチャンネル情報を登録
                var addSlackChannel = new SlackChannel()
                {
                    Name = opts.Name,
                    WebhookUrl = opts.WebhookUrl,
                    RegisteredBy = opts.RegisteredBy
                };
                var endPointUrl = $"{opts.EndpointUrl}{"SlackChannels"}";
                var postData = JsonConvert.SerializeObject(addSlackChannel);
                var postContent = new StringContent(postData, Encoding.UTF8, "application/json");
                var response = s_HttpClient.PostAsync(endPointUrl, postContent).Result;
                addSlackChannel = JsonConvert.DeserializeObject<SlackChannel>(response.Content.ReadAsStringAsync().Result);

                #endregion

                #region 登録したSlackチャンネル情報のIDと引数のWeb会議情報通知サービスのエンドポイントURLをsetting.jsonに保存

                var setting = new Setting()
                {
                    SlackChannelId = addSlackChannel.Id,
                    Name = addSlackChannel.Name,
                    RegisteredBy = addSlackChannel.RegisteredBy,
                    EndpointUrl = opts.EndpointUrl
                };

                // jsonに設定を出力
                var settingJsonString = JsonConvert.SerializeObject(setting);
                if (File.Exists(opts.Filepath))
                {
                    File.Delete(opts.Filepath);
                }

                using (var fs = File.CreateText(opts.Filepath))
                {
                    fs.WriteLine(settingJsonString);
                }

                #endregion

                return 1;
            };
            Func<RegisterOptions, int> RunRegisterAndReturnExitCode = opts =>
            {
                Console.WriteLine("Run Register");
                Console.WriteLine($"filepath:{opts.Filepath}");

                var application = new Outlook.Application();

                #region ログインユーザーのOutlookから、翌稼働日の予定を取得

                // ログインユーザーのOutlookの予定表フォルダを取得
                Outlook.Folder calFolder =
                    application.Session.GetDefaultFolder(
                            Outlook.OlDefaultFolders.olFolderCalendar)
                        as Outlook.Folder;

                DateTime startDate = DateTime.Today.AddDays(1);
                DateTime endDate = startDate.AddDays(1);
                Outlook.Items nextOperatingDayAppointments = GetAppointmentsInRange(calFolder, startDate, endDate);
                Console.WriteLine($"nextOperatingDayAppointments.Count:{nextOperatingDayAppointments.Count}");

                #endregion

                #region 取得した予定一覧の中からWeb会議情報を含む予定を抽出

                var webMeetingAppointments = new List<Outlook.AppointmentItem>();

                // ZoomURLを特定するための正規表現
                var webMeetingUrlRegexp = @"https?://[^(?!.*(/|.|\n).*$)]*\.?(zoom\.us|teams\.live\.com|teams\.microsoft\.com)/[A-Za-z0-9/?=]+";

                foreach (Outlook.AppointmentItem nextOperatingDayAppointment in nextOperatingDayAppointments)
                {
                    var appointmentBody = nextOperatingDayAppointment.Body;
                    // ZoomURLが本文に含まれる予定を正規表現で検索し、リストに詰める
                    if (!String.IsNullOrEmpty(appointmentBody) && Regex.IsMatch(appointmentBody, webMeetingUrlRegexp))
                    {
                        webMeetingAppointments.Add(nextOperatingDayAppointment);
                    }
                }
                Console.WriteLine(JsonConvert.SerializeObject(webMeetingAppointments));

                #endregion

                #region 引数のパスに存在するsetting.jsonに設定されているSlackチャンネル情報のIDと抽出した予定を使い、Web会議情報を作成

                // jsonファイルから設定を取り出す
                var fileContent = string.Empty;
                using (var sr = new StreamReader(opts.Filepath, Encoding.GetEncoding("utf-8")))
                {
                    fileContent = sr.ReadToEnd();
                }

                var setting = JsonConvert.DeserializeObject<Setting>(fileContent);

                // 追加する会議情報の一覧を作成
                var addWebMettings = new List<WebMeeting>();
                foreach (var webMeetingAppointment in webMeetingAppointments)
                {
                    var url = Regex.Match(webMeetingAppointment.Body, webMeetingUrlRegexp).Value;
                    var name = webMeetingAppointment.Subject;
                    // Outlook.AppointmentItem.Start には時刻の種類が未設定のため、現地時刻として設定する。
                    var startDateTime = new DateTime(webMeetingAppointment.Start.Ticks, DateTimeKind.Local);
                    var addWebMetting = new WebMeeting()
                    {
                        Name = name,
                        StartDateTime = startDateTime,
                        Url = url,
                        RegisteredBy = setting.RegisteredBy,
                        SlackChannelId = setting.SlackChannelId
                    };
                    Console.WriteLine(JsonConvert.SerializeObject(addWebMetting));
                    addWebMettings.Add(addWebMetting);
                }
                if(!addWebMettings.Any())
                {
                    return -1;
                }

                #endregion

                #region 引数のパスに存在するsetting.jsonに設定されているエンドポイントURLを使い、Web会議情報を削除

                var endPointUrl = $"{setting.EndpointUrl}{"WebMeetings"}";
                var getEndPointUrl = $"{endPointUrl}?fromDate={startDate}&toDate={endDate}";
                var getWebMeetingsResult = s_HttpClient.GetAsync(getEndPointUrl).Result;
                var getWebMeetingsString = getWebMeetingsResult.Content.ReadAsStringAsync().Result;
                // Getしたコンテンツはメッセージ+Jsonコンテンツなので、Jsonコンテンツだけ無理やり取り出す
                var getWebMeetings = JsonConvert.DeserializeObject<List<WebMeeting>>(getWebMeetingsString);

                foreach (var getWebMeeting in getWebMeetings)
                {
                    var deleteEndPointUrl = $"{endPointUrl}/{getWebMeeting.Id}";
                    Console.WriteLine($"[DELETE] {deleteEndPointUrl}");
                    var responseDelete = s_HttpClient.DeleteAsync(deleteEndPointUrl).Result;
                    Console.WriteLine($" responseDelete:{JsonConvert.SerializeObject(responseDelete)}");
                }

                #endregion

                #region 引数のパスに存在するsetting.jsonに設定されているエンドポイントURLを使い、Web会議情報を登録

                // Web会議情報を登録
                foreach (var addWebMetting in addWebMettings)
                {
                    var postData = JsonConvert.SerializeObject(addWebMetting);
                    var postContent = new StringContent(postData, Encoding.UTF8, "application/json");
                    Console.WriteLine($"[POST] {endPointUrl}");
                    Console.WriteLine($" Body:{postContent}");
                    var responsePost = s_HttpClient.PostAsync(endPointUrl, postContent).Result;
                    Console.WriteLine($" responseDelete:{JsonConvert.SerializeObject(responsePost)}");
                }

                #endregion

                return 1;
            };

            return CommandLine.Parser.Default.ParseArguments<SettingOptions, RegisterOptions>(args)
                .MapResult(
                    (SettingOptions opts) => RunSettingAndReturnExitCode(opts),
                    (RegisterOptions opts) => RunRegisterAndReturnExitCode(opts),
                    errs => 1);
        }

        /// <summary>
        /// 指定したOutlookフォルダから指定期間の予定を取得する
        /// </summary>
        /// <param name="folder">Outlookフォルダ</param>
        /// <param name="startTime">開始日時</param>
        /// <param name="endTime">終了日時</param>
        /// <returns>Outlook.Items</returns>
        private static Outlook.Items GetAppointmentsInRange(
            Outlook.Folder folder, DateTime startTime, DateTime endTime)
        {
            string filter = "[Start] >= '"
                            + startTime.ToString("g")
                            + "' AND [End] <= '"
                            + endTime.ToString("g") + "'";
            Console.WriteLine(filter);
            try
            {
                Outlook.Items calItems = folder.Items;
                calItems.IncludeRecurrences = true;
                calItems.Sort("[Start]", Type.Missing);
                Outlook.Items restrictItems = calItems.Restrict(filter);
                Console.WriteLine($"Outlook.Items.Count:{restrictItems.Count}");
                if (restrictItems.Count > 0)
                {
                    return restrictItems;
                }
                else
                {
                    return null;
                }
            }
            catch { return null; }
        }
    }
}
