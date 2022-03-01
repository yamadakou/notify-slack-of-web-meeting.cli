﻿using System;
using Newtonsoft.Json;

namespace notify_slack_of_web_meeting.cli.WebMeetings
{
    /// <summary>
    /// Web会議情報
    /// </summary>
    public class WebMeeting
    {
        /// <summary>
        /// 既定のコンストラクタ
        /// </summary>
        public WebMeeting()
        {
            Id = Guid.NewGuid().ToString();
        }

        /// <summary>
        /// 一意とするID
        /// </summary>
        [JsonProperty("id")]
        public string Id { get; set; }

        /// <summary>
        /// Web会議名
        /// </summary>
        [JsonProperty("name")]
        public string Name { get; set; }

        /// <summary>
        /// Web会議の開始日時
        /// </summary>
        [JsonProperty("startDateTime")]
        public DateTime StartDateTime { get; set; }

        /// <summary>
        /// Web会議のURL
        /// </summary>
        [JsonProperty("url")]
        public string Url { get; set; }

        /// <summary>
        /// 登録者
        /// </summary>
        [JsonProperty("registeredBy")]
        public string RegisteredBy { get; set; }

        /// <summary>
        /// 登録日時（UTC）
        /// </summary>
        [JsonProperty("registeredAt")]
        public DateTime RegisteredAt { get; set; }

        /// <summary>
        /// 通知先のSlackチャンネル
        /// </summary>
        [JsonProperty("slackChannelId")]
        public string SlackChannelId { get; set; }
    }
}
