// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
using System;
using Application.Common;
using Application.Common.Models;
using BotService.Infrastructure.Common;

namespace Infrastructure.Core.Common
{
    public class MeetingUrlHelper : IMeetingUrlHelper
    {
        private ParsedMeetingUrl parsedMeetingUrl;

        public void Init(string joinUrl)
        {
            parsedMeetingUrl = MeetingUrlParser.Parse(joinUrl);
        }

        public string GetThreadId()
        {
            return parsedMeetingUrl.ThreadId;
        }

        public string GetMessageId()
        {
            return parsedMeetingUrl.MessageId;
        }

        public JoinUrlContext GetContext()
        {
            return parsedMeetingUrl.Context;
        }

        public ParsedMeetingUrl GetParsedMeetingUrl()
        {
            return parsedMeetingUrl;
        }

        public string GetMeetingId()
        {
            return parsedMeetingUrl.MeetingId;
        }

        public string GetPasscode()
        {
            return parsedMeetingUrl.Passcode;
        }

        public MeetingJoinUrlType GetUrlType()
        {
            return parsedMeetingUrl.JoinUrlType;
        }
    }
}
