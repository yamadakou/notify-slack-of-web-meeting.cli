using System;
using System.Collections.Generic;
using System.Text;
using CommandLine;

namespace notify_slack_of_web_meeting.cli
{
    class Program
    {
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

                // 引数の値でSlackチャンネル情報を登録

                // 登録したSlackチャンネル情報のIDと引数のWeb会議情報通知サービスのエンドポイントURLをsetting.jsonに保存

                return 1;
            };
            Func<RegisterOptions, int> RunRegisterAndReturnExitCode = opts =>
            {
                Console.WriteLine("Run Register");
                Console.WriteLine($"filepath:{opts.Filepath}");

                // ログインユーザーのOutlookから、翌稼働日の予定を取得

                // 取得した予定一覧の中からWeb会議情報を含む予定を抽出

                // 引数のパスに存在するsetting.jsonに設定されているSlackチャンネル情報のIDと抽出した予定を使い、Web会議情報を作成

                // 引数のパスに存在するsetting.jsonに設定されているエンドポイントURLを使い、Web会議情報を削除

                // 引数のパスに存在するsetting.jsonに設定されているエンドポイントURLを使い、Web会議情報を登録

                return 1;
            };

            return CommandLine.Parser.Default.ParseArguments<SettingOptions, RegisterOptions>(args)
                .MapResult(
                    (SettingOptions opts) => RunSettingAndReturnExitCode(opts),
                    (RegisterOptions opts) => RunRegisterAndReturnExitCode(opts),
                    errs => 1);
        }
    }
}
