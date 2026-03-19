// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
namespace Application.Common.Models
{
    public class ParsedMeetingUrl
    {
        public MeetingJoinUrlType JoinUrlType { get; set; }

        public string MeetingId { get; set; }

        public string ThreadId { get; set; }

        public string MessageId { get; set; }

        public string Passcode { get; set; }

        public JoinUrlContext Context { get; set; }
    }
}