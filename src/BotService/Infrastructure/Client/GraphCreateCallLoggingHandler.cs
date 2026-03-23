using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;

namespace BotService.Infrastructure.Client
{
    internal sealed class GraphCreateCallLoggingHandler : DelegatingHandler
    {
        private readonly ILogger _logger;

        public GraphCreateCallLoggingHandler(HttpMessageHandler innerHandler, ILogger logger)
            : base(innerHandler)
        {
            _logger = logger;
        }

        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            var shouldLog = ShouldLog(request);

            if (shouldLog)
            {
                _logger.LogInformation(
                    "[GraphHttp] Sending {method} {uri}",
                    request.Method,
                    request.RequestUri);
            }

            var response = await base.SendAsync(request, cancellationToken).ConfigureAwait(false);

            if (!shouldLog)
            {
                return response;
            }

            var responseBody = response.Content == null
                ? string.Empty
                : await response.Content.ReadAsStringAsync().ConfigureAwait(false);

            _logger.LogInformation(
                "[GraphHttp] Received {statusCode} {reasonPhrase} for {method} {uri}. Headers={headers} Body={body}",
                (int)response.StatusCode,
                response.ReasonPhrase,
                request.Method,
                request.RequestUri,
                FormatHeaders(response),
                string.IsNullOrWhiteSpace(responseBody) ? "<empty>" : responseBody);

            if (response.Content != null)
            {
                var mediaType = response.Content.Headers?.ContentType?.MediaType ?? "application/json";
                var replacementContent = new StringContent(responseBody, Encoding.UTF8, mediaType);

                foreach (var header in response.Content.Headers)
                {
                    replacementContent.Headers.Remove(header.Key);
                    replacementContent.Headers.TryAddWithoutValidation(header.Key, header.Value);
                }

                response.Content = replacementContent;
            }

            return response;
        }

        private static bool ShouldLog(HttpRequestMessage request)
        {
            if (request?.RequestUri == null)
            {
                return false;
            }

            return request.Method == HttpMethod.Post
                && request.RequestUri.AbsolutePath.IndexOf("/communications/calls", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string FormatHeaders(HttpResponseMessage response)
        {
            IEnumerable<KeyValuePair<string, IEnumerable<string>>> headers = Enumerable.Empty<KeyValuePair<string, IEnumerable<string>>>();

            if (response?.Headers != null)
            {
                headers = headers.Concat(response.Headers);
            }

            if (response?.Content?.Headers != null)
            {
                headers = headers.Concat(response.Content.Headers);
            }

            return string.Join(
                "; ",
                headers.Select(header => $"{header.Key}={string.Join(",", header.Value)}"));
        }
    }
}