// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
using Application.Common.Models;

namespace BotService.Infrastructure.Common
{
    public interface IMeetingUrlHelper
    {
        ParsedMeetingUrl GetParsedMeetingUrl();

        JoinUrlContext GetContext();

        string GetMeetingId();

        string GetMessageId();

        string GetPasscode();

        string GetThreadId();

        MeetingJoinUrlType GetUrlType();

        void Init(string joinUrl);
    }
}