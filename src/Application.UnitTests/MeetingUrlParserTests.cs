using System;
using Application.Call.Commands;
using Application.Common;
using Application.Common.Models;
using Xunit;

namespace Application.UnitTests
{
    public class MeetingUrlParserTests
    {
        [Fact]
        public void Parse_ShouldSupportLegacyMeetingUrlFormat()
        {
            var joinUrl = "https://teams.microsoft.com/l/meetup-join/19:meeting_abc@thread.v2/0?context={\"Tid\":\"tenant-id\",\"Oid\":\"organizer-id\",\"MessageId\":\"reply-chain-id\"}";

            var parsedMeetingUrl = MeetingUrlParser.Parse(joinUrl);

            Assert.Equal(MeetingJoinUrlType.Legacy, parsedMeetingUrl.JoinUrlType);
            Assert.Equal("19:meeting_abc@thread.v2", parsedMeetingUrl.ThreadId);
            Assert.Equal("0", parsedMeetingUrl.MessageId);
            Assert.Equal("tenant-id", parsedMeetingUrl.Context.Tid);
            Assert.Equal("organizer-id", parsedMeetingUrl.Context.Oid);
            Assert.Equal("reply-chain-id", parsedMeetingUrl.Context.MessageId);
            Assert.Equal(Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes("0#19:meeting_abc@thread.v2#0")), parsedMeetingUrl.MeetingId);
        }

        [Fact]
        public void Parse_ShouldSupportJoinMeetingIdUrlFormat()
        {
            var joinUrl = "https://teams.microsoft.com/meet/254456896?p=ghdsiogdsu";

            var parsedMeetingUrl = MeetingUrlParser.Parse(joinUrl);

            Assert.Equal(MeetingJoinUrlType.JoinMeetingId, parsedMeetingUrl.JoinUrlType);
            Assert.Equal("254456896", parsedMeetingUrl.MeetingId);
            Assert.Equal("ghdsiogdsu", parsedMeetingUrl.Passcode);
            Assert.Null(parsedMeetingUrl.Context);
            Assert.Null(parsedMeetingUrl.ThreadId);
            Assert.Null(parsedMeetingUrl.MessageId);
        }

        [Fact]
        public void Validator_ShouldAcceptJoinMeetingIdUrlFormat()
        {
            var validator = new RequestInviteBot.RequestInviteBotCommandValidator();
            var command = new RequestInviteBot.RequestInviteBotCommand
            {
                MeetingUrl = "https://teams.microsoft.com/meet/254456896?p=ghdsiogdsu",
            };

            var validationResult = validator.Validate(command);

            Assert.True(validationResult.IsValid);
        }

        [Fact]
        public void Validator_ShouldRejectInvalidMeetingUrlFormat()
        {
            var validator = new RequestInviteBot.RequestInviteBotCommandValidator();
            var command = new RequestInviteBot.RequestInviteBotCommand
            {
                MeetingUrl = "https://teams.microsoft.com/meet/254456896",
            };

            var validationResult = validator.Validate(command);

            Assert.False(validationResult.IsValid);
        }
    }
}