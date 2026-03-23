using System.Linq;
using Application.Common.Notifications;
using Xunit;

namespace Application.UnitTests
{
    public class CallingNotificationPayloadParserTests
    {
        [Fact]
        public void Parse_ShouldReturnEstablishedAndParticipantsNotifications()
        {
            const string payload = @"{
  ""@odata.type"": ""#microsoft.graph.commsNotifications"",
  ""value"": [
    {
      ""@odata.type"": ""#microsoft.graph.commsNotification"",
      ""changeType"": ""updated"",
      ""resource"": ""/app/calls/05003180-2276-42f6-a3f3-7a032907b3c1"",
      ""resourceData"": {
        ""@odata.type"": ""#microsoft.graph.call"",
        ""state"": ""established""
      }
    },
    {
      ""@odata.type"": ""#microsoft.graph.commsNotification"",
      ""changeType"": ""updated"",
      ""resource"": ""/app/calls/05003180-2276-42f6-a3f3-7a032907b3c1/participants"",
      ""resourceData"": [
        {
          ""@odata.type"": ""#microsoft.graph.participant"",
          ""id"": ""516b0047-69f0-4c5a-8148-e29f41c540d1"",
          ""isMuted"": true,
          ""mediaStreams"": [
            {
              ""mediaType"": ""audio"",
              ""direction"": ""sendReceive""
            }
          ]
        }
      ]
    }
  ]
}";

            var notifications = CallingNotificationPayloadParser.Parse(payload);

            Assert.Equal(2, notifications.Count);
            Assert.All(notifications, notification => Assert.Equal("05003180-2276-42f6-a3f3-7a032907b3c1", notification.GraphCallId));

            var establishedNotification = notifications.Single(x => !x.IsParticipantsNotification);
            Assert.Equal("established", establishedNotification.ResourceData.Value<string>("state"));

            var participantsNotification = notifications.Single(x => x.IsParticipantsNotification);
            Assert.True(participantsNotification.ResourceData.HasValues);
        }
    }
}