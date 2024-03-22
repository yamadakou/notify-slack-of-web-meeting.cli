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
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Http;
using Polly;
using Polly.Extensions.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using NLog;
using NLog.Extensions.Logging;

// # ※リトライ可能なHTTP要求を実現するための参考
// ## Microsoft Docs
// * IHttpClientFactory を使用して回復力の高い HTTP 要求を実装する
//   * https://docs.microsoft.com/ja-jp/dotnet/architecture/microservices/implement-resilient-applications/use-httpclientfactory-to-implement-resilient-http-requests
// * IHttpClientFactory ポリシーと Polly ポリシーで指数バックオフを含む HTTP 呼び出しの再試行を実装する
//   * https://docs.microsoft.com/ja-jp/dotnet/architecture/microservices/implement-resilient-applications/implement-http-call-retries-exponential-backoff-polly
// ## Blog
// * C# - HttpClientFactoryをDIの外で使う
//   * https://dekirukigasuru.com/blog/2020/04/24/csharp-httpclientfactory/
// * C# - HttpClientFactoryとPollyで回復力の高い何某
//   * https://dekirukigasuru.com/blog/2020/05/15/csharp-httpclientfactory-polly/
// ## Nuget
// * Microsoft.Extensions.DependencyInjection
//   * https://www.nuget.org/packages/Microsoft.Extensions.DependencyInjection/5.0.2
// * Microsoft.Extensions.Http
//   * https://www.nuget.org/packages/Microsoft.Extensions.Http/5.0.0
// * Microsoft.Extensions.Http.Polly
//   * https://www.nuget.org/packages/Microsoft.Extensions.Http.Polly/5.0.1

// # ※NLogの参考
// ## チュートリアル
// * https://github.com/NLog/NLog/wiki/Getting-started-with-.NET-Core-2---Console-application


namespace notify_slack_of_web_meeting.cli
{
    class Program
    {
        /// <summary>
        /// Logger
        /// </summary>
        private static Logger s_logger = LogManager.GetCurrentClassLogger();

        /// <summary>
        /// HTTPクライアント
        /// </summary>
        private static HttpClient s_HttpClient;

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

            [Option('d', "days", HelpText = "Number of days to get an appointment.", Default = 1)]
            public int Days { get; set; }
        }

        static int Main(string[] args)
        {
            #region HTTPクライアントを設定

            // Serviceにリトライ可能なHTTPクライアントを設定
            var services = new ServiceCollection();
            services.AddHttpClient("RetryHttpClient")
            .SetHandlerLifetime(TimeSpan.FromMinutes(5))  // ライフタイムを5分に設定
            .AddPolicyHandler(GetRetryPolicy());

            // HTTPクライアントファクトリーを取得
            var factory = services.BuildServiceProvider().GetRequiredService<IHttpClientFactory>();

            // リトライ可能なHTTPクライアントを取得
            s_HttpClient = factory.CreateClient("RetryHttpClient");

            #endregion

            #region Settingコマンド

            // Settingコマンドを定義
            Func<SettingOptions, int> RunSettingAndReturnExitCode = opts =>
            {
                s_logger.Info("Run Setting");

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
                var settingJsonString = JsonConvert.SerializeObject(setting, Formatting.Indented);
                s_logger.Info(settingJsonString);
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

            #endregion

            #region  Registerコマンド

            // Registerコマンドを定義
            Func<RegisterOptions, int> RunRegisterAndReturnExitCode = opts =>
            {
                s_logger.Info("Run Register");
                s_logger.Debug($"filepath:{opts.Filepath}");
                s_logger.Debug($"days:{opts.Days}");

                var application = new Outlook.Application();

                #region ログインユーザーのOutlookから、翌稼働日の予定を取得

                // ログインユーザーのOutlookの予定表フォルダを取得
                Outlook.Folder calFolder =
                    application.Session.GetDefaultFolder(
                            Outlook.OlDefaultFolders.olFolderCalendar)
                        as Outlook.Folder;

                int days = opts.Days < 1 ? 1 : opts.Days;
                DateTime startDate = DateTime.Today.AddDays(1);
                DateTime endDate = startDate.AddDays(days);
                Outlook.Items nextOperatingDayAppointments = GetAppointmentsInRange(calFolder, startDate, endDate);
                s_logger.Debug($"nextOperatingDayAppointments.Count:{nextOperatingDayAppointments.Count}");

                #endregion

                #region 取得した予定一覧の中からWeb会議情報を含む予定を抽出

                var webMeetingAppointments = new List<Outlook.AppointmentItem>();

                // Web会議（Zoom/Teams）のURLを特定するための正規表現
                var webMeetingUrlRegexp = @"https?(://|%3A%2F%2F)[^(?!.*(/|.|\n).*$)]*\.?(zoom\.us|teams\.live\.com|teams\.microsoft\.com)(/|%2F)[%A-Za-z0-9/?=]+";

                foreach (Outlook.AppointmentItem nextOperatingDayAppointment in nextOperatingDayAppointments)
                {
                    var appointmentBody = nextOperatingDayAppointment.Body;
                    // Web会議のURLが本文に含まれる予定を正規表現で検索し、リストに詰める
                    if (!String.IsNullOrEmpty(appointmentBody) && Regex.IsMatch(appointmentBody, webMeetingUrlRegexp))
                    {
                        webMeetingAppointments.Add(nextOperatingDayAppointment);
                    }
                }
                s_logger.Debug(JsonConvert.SerializeObject(webMeetingAppointments));

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
                    // Web会議のURLがURIエンコードされているれる場合を考慮し、URIデコードしてURLとして設定する。
                    var url = Uri.UnescapeDataString(Regex.Match(webMeetingAppointment.Body, webMeetingUrlRegexp).Value);
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
                   s_logger.Debug(JsonConvert.SerializeObject(addWebMetting));
                    addWebMettings.Add(addWebMetting);
                }
                if(!addWebMettings.Any())
                {
                    return -1;
                }

                #endregion

                #region 引数のパスに存在するsetting.jsonに設定されているエンドポイントURLを使い、Web会議情報を削除

                var endPointUrl = $"{setting.EndpointUrl}{"WebMeetings"}";
                var getEndPointUrl = $"{endPointUrl}?fromDate={startDate}&slackChannelId={setting.SlackChannelId}";
                var getWebMeetingsResult = s_HttpClient.GetAsync(getEndPointUrl).Result;
                var getWebMeetingsString = getWebMeetingsResult.Content.ReadAsStringAsync().Result;
                // Getしたコンテンツはメッセージ+Jsonコンテンツなので、Jsonコンテンツだけ無理やり取り出す
                var getWebMeetings = JsonConvert.DeserializeObject<List<WebMeeting>>(getWebMeetingsString);

                foreach (var getWebMeeting in getWebMeetings)
                {
                    var deleteEndPointUrl = $"{endPointUrl}/{getWebMeeting.Id}";
                   s_logger.Info($"[DELETE] {deleteEndPointUrl}");
                    var responseDelete = s_HttpClient.DeleteAsync(deleteEndPointUrl).Result;
                   s_logger.Info($" responseDelete:{JsonConvert.SerializeObject(responseDelete)}");
                }

                #endregion

                #region 引数のパスに存在するsetting.jsonに設定されているエンドポイントURLを使い、Web会議情報を登録

                // Web会議情報を登録
                foreach (var addWebMetting in addWebMettings)
                {
                    var postData = JsonConvert.SerializeObject(addWebMetting);
                    var postContent = new StringContent(postData, Encoding.UTF8, "application/json");
                   s_logger.Info($"[POST] {endPointUrl}");
                   s_logger.Info($" Body: {JsonConvert.SerializeObject(addWebMetting, Formatting.Indented)}");
                    var responsePost = s_HttpClient.PostAsync(endPointUrl, postContent).Result;
                   s_logger.Info($" responsePost:{JsonConvert.SerializeObject(responsePost)}");
                }

                #endregion

                return 1;
            };

            #endregion

            // コマンドを実行
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
           s_logger.Debug(filter);
            try
            {
                Outlook.Items calItems = folder.Items;
                calItems.IncludeRecurrences = true;
                calItems.Sort("[Start]", Type.Missing);
                Outlook.Items restrictItems = calItems.Restrict(filter);
               s_logger.Debug($"Outlook.Items.Count:{restrictItems.Count}");
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

        /// <summary>
        /// リトライのポリシーを取得する
        /// </summary>
        /// <returns>リトライのポリシー</returns>
        private static IAsyncPolicy<HttpResponseMessage> GetRetryPolicy()
        {
            // 指数関数的再試行で (最初は 2 秒) 6 回試すポリシー
            // リトライ対象となる条件は以下
            // * Network failures (as HttpRequestException)
            // * HTTP 5XX status codes (server errors)
            // * HTTP 408 status code (request timeout)
            // * HTTP 429 status code (too many requests)
            // ※参考
            // https://github.com/App-vNext/Polly/wiki/Polly-and-HttpClientFactory#using-addtransienthttperrorpolicy
            // https://github.com/App-vNext/Polly.Extensions.Http/blob/master/src/Polly.Extensions.Http/HttpPolicyExtensions.cs#L17-L28
            return HttpPolicyExtensions
                .HandleTransientHttpError()
                .OrResult(msg => msg.StatusCode == System.Net.HttpStatusCode.TooManyRequests)
                .WaitAndRetryAsync(6, retryAttempt => TimeSpan.FromSeconds(Math.Pow(2, retryAttempt)));
        }

    }
}
