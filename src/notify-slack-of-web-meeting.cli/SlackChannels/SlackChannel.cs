﻿using System;
using Newtonsoft.Json;

namespace notify_slack_of_web_meeting.cli.SlackChannels
{
    /// <summary>
    /// Slackチャンネル
    /// </summary>
    public class SlackChannel
    {
        /// <summary>
        /// 一意とするID
        /// </summary>
        [JsonProperty("id")]
        public string Id { get; set; }

        /// <summary>
        /// Slackチャンネル名
        /// </summary>
        [JsonProperty("name")]
        public string Name { get; set; }

        /// <summary>
        /// SlackチャンネルのWebhook URL
        /// </summary>
        [JsonProperty("webhookUrl")]
        public string WebhookUrl { get; set; }

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
    }
}
