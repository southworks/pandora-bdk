// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using Application.Common.Models;

namespace Application.Common
{
    public static class MeetingUrlParser
    {
        public static ParsedMeetingUrl Parse(string joinUrl)
        {
            if (!TryParse(joinUrl, out var parsedMeetingUrl))
            {
                throw new ArgumentException($"Join URL cannot be parsed: {joinUrl}.", nameof(joinUrl));
            }

            return parsedMeetingUrl;
        }

        public static bool TryParse(string joinUrl, out ParsedMeetingUrl parsedMeetingUrl)
        {
            parsedMeetingUrl = null;

            if (string.IsNullOrWhiteSpace(joinUrl))
            {
                return false;
            }

            var decodedUrl = WebUtility.UrlDecode(joinUrl);

            if (!Uri.TryCreate(decodedUrl, UriKind.Absolute, out var uri))
            {
                return false;
            }

            if (!uri.Host.Equals("teams.microsoft.com", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            var pathSegments = uri.AbsolutePath.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            var queryParameters = ParseQuery(uri.Query);

            if (TryParseJoinMeetingIdUrl(pathSegments, queryParameters, out parsedMeetingUrl))
            {
                return true;
            }

            if (TryParseLegacyUrl(pathSegments, queryParameters, out parsedMeetingUrl))
            {
                return true;
            }

            return false;
        }

        private static bool TryParseJoinMeetingIdUrl(IReadOnlyList<string> pathSegments, IReadOnlyDictionary<string, string> queryParameters, out ParsedMeetingUrl parsedMeetingUrl)
        {
            parsedMeetingUrl = null;

            if (pathSegments.Count < 2 || !pathSegments[pathSegments.Count - 2].Equals("meet", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            if (!queryParameters.TryGetValue("p", out var passcode) || string.IsNullOrWhiteSpace(passcode))
            {
                return false;
            }

            var joinMeetingId = pathSegments[pathSegments.Count - 1];
            if (string.IsNullOrWhiteSpace(joinMeetingId))
            {
                return false;
            }

            parsedMeetingUrl = new ParsedMeetingUrl
            {
                JoinUrlType = MeetingJoinUrlType.JoinMeetingId,
                MeetingId = joinMeetingId,
                Passcode = passcode,
            };

            return true;
        }

        private static bool TryParseLegacyUrl(IReadOnlyList<string> pathSegments, IReadOnlyDictionary<string, string> queryParameters, out ParsedMeetingUrl parsedMeetingUrl)
        {
            parsedMeetingUrl = null;

            if (pathSegments.Count < 2 || !queryParameters.TryGetValue("context", out var contextJson) || string.IsNullOrWhiteSpace(contextJson))
            {
                return false;
            }

            if (!TryDeserializeContext(contextJson, out var context))
            {
                return false;
            }

            var threadId = pathSegments[pathSegments.Count - 2];
            var messageId = pathSegments[pathSegments.Count - 1];

            if (string.IsNullOrWhiteSpace(threadId) || string.IsNullOrWhiteSpace(messageId))
            {
                return false;
            }

            parsedMeetingUrl = new ParsedMeetingUrl
            {
                JoinUrlType = MeetingJoinUrlType.Legacy,
                ThreadId = threadId,
                MessageId = messageId,
                Context = context,
                MeetingId = Base64Encode($"0#{threadId}#0"),
            };

            return true;
        }

        private static Dictionary<string, string> ParseQuery(string query)
        {
            var parameters = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var trimmedQuery = query?.TrimStart('?');

            if (string.IsNullOrWhiteSpace(trimmedQuery))
            {
                return parameters;
            }

            var pairs = trimmedQuery.Split(new[] { '&' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var pair in pairs)
            {
                var parts = pair.Split(new[] { '=' }, 2);
                var key = WebUtility.UrlDecode(parts[0]);
                var value = parts.Length > 1 ? WebUtility.UrlDecode(parts[1]) : string.Empty;

                parameters[key] = value;
            }

            return parameters;
        }

        private static bool TryDeserializeContext(string contextJson, out JoinUrlContext context)
        {
            context = null;

            try
            {
                using var stream = new MemoryStream(Encoding.UTF8.GetBytes(contextJson));
                var serializer = new DataContractJsonSerializer(typeof(JoinUrlContext));
                context = serializer.ReadObject(stream) as JoinUrlContext;
                return context != null;
            }
            catch (SerializationException)
            {
                return false;
            }
        }

        private static string Base64Encode(string plainText)
        {
            var plainTextBytes = Encoding.UTF8.GetBytes(plainText);
            return Convert.ToBase64String(plainTextBytes);
        }
    }
}