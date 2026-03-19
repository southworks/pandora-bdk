// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
using System.Collections.Generic;
using Microsoft.Graph;

namespace BotService.Infrastructure.Common
{
    public class JoinMeetingIdMeetingInfo : MeetingInfo
    {
        public JoinMeetingIdMeetingInfo(string joinMeetingId, string passcode)
        {
            ODataType = "#microsoft.graph.joinMeetingIdMeetingInfo";
            AdditionalData = new Dictionary<string, object>
            {
                { "joinMeetingId", joinMeetingId },
                { "passcode", passcode },
            };
        }
    }
}