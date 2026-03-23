// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
using System.Collections.Generic;

namespace Microsoft.Graph
{
    public class JoinMeetingIdMeetingInfo : MeetingInfo
    {
        public JoinMeetingIdMeetingInfo()
        {
            ODataType = "#microsoft.graph.joinMeetingIdMeetingInfo";
            AdditionalData = new Dictionary<string, object>();
        }

        public JoinMeetingIdMeetingInfo(string joinMeetingId, string passcode)
            : this()
        {
            AdditionalData["joinMeetingId"] = joinMeetingId;
            AdditionalData["passcode"] = passcode;
        }
    }
}