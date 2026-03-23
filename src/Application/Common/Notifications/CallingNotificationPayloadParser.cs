using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;

namespace Application.Common.Notifications
{
    public class ParsedCallingNotification
    {
        public string Resource { get; set; }

        public string GraphCallId { get; set; }

        public JToken ResourceData { get; set; }

        public bool IsParticipantsNotification =>
            !string.IsNullOrWhiteSpace(Resource)
            && Resource.EndsWith("/participants", StringComparison.OrdinalIgnoreCase);
    }

    public static class CallingNotificationPayloadParser
    {
        public static IReadOnlyList<ParsedCallingNotification> Parse(string payload)
        {
            if (string.IsNullOrWhiteSpace(payload))
            {
                return Array.Empty<ParsedCallingNotification>();
            }

            JToken rootToken;
            try
            {
                rootToken = JToken.Parse(payload);
            }
            catch
            {
                return Array.Empty<ParsedCallingNotification>();
            }

            if (!(rootToken["value"] is JArray notifications))
            {
                return Array.Empty<ParsedCallingNotification>();
            }

            return notifications
                .OfType<JObject>()
                .Select(notification =>
                {
                    var resource = notification.Value<string>("resource");
                    return new ParsedCallingNotification
                    {
                        Resource = resource,
                        GraphCallId = TryGetGraphCallId(resource, out var graphCallId) ? graphCallId : null,
                        ResourceData = notification["resourceData"],
                    };
                })
                .Where(x => !string.IsNullOrWhiteSpace(x.GraphCallId) && x.ResourceData != null)
                .ToList();
        }

        private static bool TryGetGraphCallId(string resource, out string graphCallId)
        {
            graphCallId = null;

            if (string.IsNullOrWhiteSpace(resource))
            {
                return false;
            }

            var segments = resource.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            for (int index = 0; index < segments.Length - 1; index++)
            {
                if (!string.Equals(segments[index], "calls", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                graphCallId = segments[index + 1];
                return !string.IsNullOrWhiteSpace(graphCallId);
            }

            return false;
        }
    }
}