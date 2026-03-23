// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Application.Call.Specifications;
using Application.Common.Config;
using Application.Common.Models;
using Application.Common.Models.Api;
using Application.Common.Notifications;
using Application.Interfaces.Common;
using Application.Interfaces.Persistance;
using Application.Participants.Specifications;
using Application.Service.Commands;
using BotService.Application.Core;
using BotService.Infrastructure.Common;
using BotService.Infrastructure.Extensions;
using BotService.Infrastructure.Pipelines;
using BotService.Infrastructure.Services;
using Domain.Entities;
using Domain.Entities.Parts;
using Domain.Enums;
using Domain.Exceptions;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Communications.Calls;
using Microsoft.Graph.Communications.Calls.Media;
using Microsoft.Graph.Communications.Client;
using Microsoft.Graph.Communications.Common;
using Microsoft.Graph.Communications.Common.Telemetry;
using Microsoft.Graph.Communications.Resources;
using Microsoft.Skype.Bots.Media;
using Newtonsoft.Json.Linq;
using Serilog;

namespace BotService.Infrastructure.Core
{
    public class Bot : IBot, IDisposable
    {
        private readonly ICommunicationsClient _client;
        private readonly IMediatorService _mediatorService;
        private readonly IMediaHandlerFactory _mediaHandlerFactory;
        private readonly ILoggerFactory _loggerFactory;
        private readonly BotConfiguration _config;
        private readonly string _tenantId;
        private readonly ILogger<Bot> _logger;
        private readonly GstreamerClockProvider _clockProvider;
        private readonly IServiceScopeFactory _serviceScopeFactory;
        private string _activeScenarioCallId;

        public Bot(
            ICommunicationsClient client,
            IMediatorService mediatorService,
            IMediaHandlerFactory mediaHandlerFactory,
            IAppConfiguration config,
            GstreamerClockProvider clockProvider,
            ILoggerFactory loggerFactory,
            IServiceScopeFactory serviceScopeFactory)
        {
            _client = client;
            _mediatorService = mediatorService;
            _mediaHandlerFactory = mediaHandlerFactory;
            _clockProvider = clockProvider;
            _loggerFactory = loggerFactory;
            _serviceScopeFactory = serviceScopeFactory;

            _config = config.BotConfiguration;
            _tenantId = config.AzureAdConfiguration.TenantId;
            _logger = loggerFactory.CreateLogger<Bot>();

            _client.Calls().OnIncoming += CallsOnIncoming;
            _client.Calls().OnUpdated += CallsOnUpdated;

            _logger.LogInformation("[Bot] Subscribed to ICallCollection events: OnIncoming and OnUpdated.");
        }

        /// <summary>
        /// Gets the collection of call handlers.
        /// </summary>
        public ConcurrentDictionary<string, CallHandler> CallHandlers { get; } = new ConcurrentDictionary<string, CallHandler>();

        public string Id { get; private set; }

        public string VirtualMachineName { get; set; }

        public async Task InviteBotAsync(DoInviteBot.DoInviteBotCommand command)
        {
            _logger.LogInformation("[Bot] Getting meeting info for call {callId}", command.CallId);

            MeetingInfo meetingInfo;
            ChatInfo chatInfo;

            (chatInfo, meetingInfo) = JoinInfoHelper.ParseJoinURL(command.MeetingUrl);

            var tenantId = (meetingInfo as OrganizerMeetingInfo)?.Organizer.GetPrimaryIdentity()?.GetTenantId() ?? _tenantId;
            var mediaSession = CreateLocalMediaSession();

            var joinParams = new JoinMeetingParameters(chatInfo, meetingInfo, mediaSession)
            {
                TenantId = tenantId,
            };

            var scenarioId = Guid.Parse(command.CallId);
            _activeScenarioCallId = command.CallId;

            _logger.LogInformation("[Bot] Initiating call {callId} with scenario id {scenarioId}", command.CallId, scenarioId);
            LogGraphModelResolution("Microsoft.Graph.JoinMeetingIdMeetingInfo");
            LogGraphModelResolution("Microsoft.Graph.MeetingCapability");
            LogGraphModelResolution("Microsoft.Graph.StagingRoomLiveState");

            try
            {
                var statefulCall = await _client.Calls().AddAsync(joinParams, scenarioId).ConfigureAwait(false);

                _logger.LogInformation("[Bot] Call initialization completed with response - call {callId} id={id}", command.CallId, statefulCall.Id);
                statefulCall.GraphLogger.Info($"Call creation complete: {statefulCall.Id}");
                _logger.LogInformation("[Bot] Call initialization completed - call {callId}", command.CallId);
            }
            catch (ServiceException ex)
            {
                var responseHeaders = FormatResponseHeaders(ex.ResponseHeaders);

                _logger.LogInformation("[Bot] GRAPH RESPONSE DUMP - Exception Details:");
                _logger.LogInformation($"[Bot]   HTTP StatusCode: {ex.StatusCode}");
                _logger.LogInformation($"[Bot]   Error Code: {ex.Error?.Code}");
                _logger.LogInformation($"[Bot]   Error Message: {ex.Error?.Message}");
                _logger.LogInformation($"[Bot]   RawResponseBody: {ex.RawResponseBody}");
                if (ex.ResponseHeaders != null)
                {
                    foreach (var header in ex.ResponseHeaders)
                    {
                        if (header.Key.Equals("Location", StringComparison.OrdinalIgnoreCase))
                        {
                            _logger.LogInformation($"[Bot]   Location Header: {string.Join(", ", header.Value)}");
                        }
                    }
                }

                _logger.LogInformation($"[Bot]   InnerException Type: {ex.InnerException?.GetType().FullName}");
                _logger.LogInformation($"[Bot]   InnerException Message: {ex.InnerException?.Message}");

                if (ex.InnerException != null)
                {
                    _logger.LogDebug("[Bot]   InnerException StackTrace: {stack}", ex.InnerException.StackTrace);
                }

                var isKnownDeserializationIssue =
                    ex.StatusCode == 0
                    && string.Equals(ex.Error?.Code, "generalException", StringComparison.OrdinalIgnoreCase)
                    && ex.InnerException is MissingMethodException;

                if (isKnownDeserializationIssue)
                {
                    _logger.LogWarning(
                        ex,
                        "[Bot] Graph AddAsync returned known SDK deserialization failure for call {callId} | ScenarioId={scenarioId}. Request may have succeeded; monitoring CallsOnUpdated events.",
                        command.CallId,
                        scenarioId);
                }
                else
                {
                    _logger.LogError(
                        ex,
                        "[Bot] Graph AddAsync failed for call {callId} | ScenarioId={scenarioId} | StatusCode={statusCode} | ErrorCode={errorCode} | ErrorMessage={errorMessage} | RawResponseBody={rawResponseBody} | ResponseHeaders={responseHeaders}",
                        command.CallId,
                        scenarioId,
                        ex.StatusCode,
                        ex.Error?.Code,
                        ex.Error?.Message,
                        ex.RawResponseBody,
                        responseHeaders);

                    throw;
                }
            }
            catch (Exception ex)
            {
                // The SDK may throw a deserialization exception (e.g. MissingMethodException from
                // ODataJsonConverter trying to instantiate an abstract type) after Graph has already
                // accepted the call request. The call state arrives via CallsOnUpdated webhooks so
                // it is safe to log the warning and continue rather than surfacing a false HTTP 500.
                _logger.LogWarning(
                    ex,
                    "[Bot] Non-Graph exception during AddAsync response parsing for call {callId} | ScenarioId={scenarioId} | Exception Type: {exType} | Message: {exMsg}",
                    command.CallId,
                    scenarioId,
                    ex.GetType().FullName,
                    ex.Message);
            }

            // TODO: Analyze if we need to return something
        }

        public async Task ProcessNotificationAsync(HttpRequestMessage request)
        {
            var payload = request?.Content == null
                ? string.Empty
                : await request.Content.ReadAsStringAsync().ConfigureAwait(false);

            var payloadPreview = string.Empty;
            if (!string.IsNullOrWhiteSpace(payload))
            {
                payloadPreview = payload;
                if (payload.Length > 1200)
                {
                    payloadPreview = payload.Substring(0, 1200);
                }
            }

            Log.Information("[Bot] Entering ProcessNotificationAsync. PayloadLength={payloadLength} RequestUri={requestUri} PayloadPreview={payloadPreview}", payload?.Length ?? 0, request?.RequestUri?.ToString(), payloadPreview);

            if (!string.IsNullOrWhiteSpace(payload))
            {
                try
                {
                    var payloadJson = JToken.Parse(payload);
                    var allNodes = payloadJson is JContainer container
                        ? container.DescendantsAndSelf()
                        : new[] { payloadJson };

                    var odataTypes = allNodes
                        .OfType<JObject>()
                        .Select(node => node.Value<string>("@odata.type"))
                        .Where(value => !string.IsNullOrWhiteSpace(value))
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
                        .ToArray();

                    Log.Information("[Bot] Notification payload OData types: {odataTypes}", string.Join(", ", odataTypes));
                }
                catch (Exception ex)
                {
                    Log.Warning(ex, "[Bot] Failed to extract OData types from notification payload.");
                }
            }

            Exception sdkProcessingException = null;

            try
            {
                await _client.ProcessNotificationAsync(CloneRequestMessage(request, payload)).ConfigureAwait(false);
                _logger.LogInformation("[Bot] Graph SDK notification processing completed.");
                Log.Information("[Bot] Graph SDK notification processing completed.");
            }
            catch (Exception ex)
            {
                sdkProcessingException = ex;
                _logger.LogWarning(ex, "[Bot] Graph SDK notification processing failed. CallsOnUpdated may not fire for this payload; attempting raw webhook reconciliation.");
                Log.Warning(ex, "[Bot] Graph SDK notification processing failed. CallsOnUpdated may not fire for this payload; attempting raw webhook reconciliation.");
            }

            var reconciled = await ReconcileNotificationPayloadAsync(payload).ConfigureAwait(false);
            Log.Information("[Bot] ReconcileNotificationPayloadAsync completed. Reconciled={reconciled}", reconciled);

            if (sdkProcessingException != null && !reconciled)
            {
                throw sdkProcessingException;
            }
        }

        public async Task RegisterServiceAsync(string virtualMachineName)
        {
            var response = await _mediatorService.RegisterServiceAsync(virtualMachineName);
            Id = response.Id;
        }

        public async Task UnregisterServiceAsync(string virtualMachineName)
        {
            var response = await _mediatorService.UnregisterServiceAsync(virtualMachineName);
            Id = response.Id;
        }

        public async Task MuteBotAsync(string callId)
        {
            var callHandler = ResolveCallHandlerOrThrow(callId);
            await callHandler.Call.MuteAsync();
        }

        public async Task UnmuteBotAsync(string callId)
        {
            var callHandler = ResolveCallHandlerOrThrow(callId);
            await callHandler.Call.UnmuteAsync();
        }

        public async Task EndCallAsync(string callGraphId)
        {
            try
            {
                var callHandler = GetHandlerOrThrow(callGraphId);
                callHandler.StopActiveStreams();
                await callHandler.Call.DeleteAsync().ConfigureAwait(false);
            }
            catch (Exception)
            {
                // Manually remove the call from SDK state.
                // This will trigger the ICallCollection.OnUpdated event with the removed resource.
                _client.Calls().TryForceRemove(callGraphId, out _);
            }
        }

        // TODO: Define Output
        public void StartInjection(StartStreamInjectionBody startStreamInjectionBody)
        {
            var callHandler = ResolveCallHandlerOrThrow(startStreamInjectionBody?.CallId);
            callHandler.StartInjection(startStreamInjectionBody);
        }

        public void StopInjection(string callId)
        {
            var callHandler = ResolveCallHandlerOrThrow(callId);
            callHandler.StopInjection();
        }

        public void DisplayInjection(string callId)
        {
            var callHandler = ResolveCallHandlerOrThrow(callId);
            callHandler.DisplayInjection();
        }

        public void HideInjection(string callId)
        {
            var callHandler = ResolveCallHandlerOrThrow(callId);
            callHandler.HideInjection();
        }

        public StartStreamExtractionResponse StartExtraction(StartStreamExtractionBody streamBody)
        {
            var callHandler = ResolveCallHandlerOrThrow(streamBody?.CallId);
            return callHandler.StartExtraction(streamBody);
        }

        public void StopExtraction(StopStreamExtractionBody streamBody)
        {
            var callHandler = ResolveCallHandlerOrThrow(streamBody?.CallId);
            callHandler.StopExtraction(streamBody);
        }

        public void SetInjectionVolume(string callId, SetInjectionVolumeRequest injectionVolumeRequest)
        {
            var injectionVolume = new StreamVolume
            {
                Format = injectionVolumeRequest.Format,
                Value = injectionVolumeRequest.Value,
            };

            var callHandler = ResolveCallHandlerOrThrow(callId);
            callHandler.SetInjectionVolume(injectionVolume);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                // Release the resources used by the Media Platform
                MediaPlatform.Shutdown();
            }
        }

        #region Private

        /// <summary>
        /// Creates the local media session.
        /// </summary>
        /// <param name="mediaSessionId">
        /// The media session identifier.
        /// This should be a unique value for each call.
        /// </param>
        /// <returns>The <see cref="ILocalMediaSession"/>.</returns>
        private ILocalMediaSession CreateLocalMediaSession(Guid mediaSessionId = default)
        {
            var videoSocketSettings = new List<VideoSocketSettings>
            {
                new VideoSocketSettings
                {
                    StreamDirections = StreamDirection.Sendrecv,
                    ReceiveColorFormat = VideoColorFormat.H264,
                    SupportedSendVideoFormats = new List<VideoFormat>
                    {
                        VideoFormat.NV12_1280x720_30Fps,
                        VideoFormat.NV12_1920x1080_30Fps,
                        VideoFormat.NV12_1920x1080_1_875Fps,
                    },
                    MaxConcurrentSendStreams = 1,
                },
            };

            // create the receive only sockets settings for the multiview support
            for (int i = 0; i < _config.NumberOfMultiviewSockets; i++)
            {
                videoSocketSettings.Add(new VideoSocketSettings
                {
                    StreamDirections = StreamDirection.Recvonly,
                    ReceiveColorFormat = VideoColorFormat.H264,
                });
            }

            // Create the VBSS socket settings
            var vbssSocketSettings = new VideoSocketSettings
            {
                StreamDirections = StreamDirection.Recvonly,
                ReceiveColorFormat = VideoColorFormat.H264,
                MediaType = MediaType.Vbss,
                SupportedSendVideoFormats = new List<VideoFormat>
                {
                    // fps 1.875 is required for h264 in vbss scenario.
                    VideoFormat.H264_1920x1080_1_875Fps,
                },
            };

            // create media session object, this is needed to establish call connections
            var mediaSession = _client.CreateMediaSession(
                new AudioSocketSettings
                {
                    StreamDirections = StreamDirection.Sendrecv,
                    SupportedAudioFormat = AudioFormat.Pcm16K,
                },
                videoSocketSettings,
                vbssSocketSettings,
                mediaSessionId: mediaSessionId);
            return mediaSession;
        }

        /// <summary>
        /// Incoming call handler.
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="args">The <see cref="CollectionEventArgs{TEntity}"/> instance containing the event data.</param>
        private void CallsOnIncoming(ICallCollection sender, CollectionEventArgs<ICall> args)
        {
            args.AddedResources.ForEach(call =>
            {
                IMediaSession mediaSession = Guid.TryParse(call.Id, out Guid callId)
                    ? CreateLocalMediaSession(callId)
                    : CreateLocalMediaSession();

                // Answer call
                call?.AnswerAsync(mediaSession).ForgetAndLogExceptionAsync(
                    call.GraphLogger,
                    $"Answering call {call.Id} with scenario {call.ScenarioId}.");
            });
        }

        /// <summary>
        /// Updated call handler.
        /// </summary>
        /// <param name="sender">The <see cref="ICallCollection"/> sender.</param>
        /// <param name="args">The <see cref="CollectionEventArgs{ICall}"/> instance containing the event data.</param>
        private void CallsOnUpdated(ICallCollection sender, CollectionEventArgs<ICall> args)
        {
            _logger.LogInformation("[Bot] CallsOnUpdated fired with {added} added and {removed} removed calls.", args.AddedResources?.Count ?? 0, args.RemovedResources?.Count ?? 0);

            foreach (var call in args.AddedResources)
            {
                var callHandler = new CallHandler(call, _config, _mediatorService, _mediaHandlerFactory, _loggerFactory, _clockProvider);
                CallHandlers[call.Id] = callHandler;
                _logger.LogInformation("[Bot] Registered CallHandler for graph call {graphCallId} and scenario call {scenarioCallId}.", call.Id, call.ScenarioId);
            }

            foreach (var call in args.RemovedResources)
            {
                if (CallHandlers.TryRemove(call.Id, out CallHandler handler))
                {
                    handler.Dispose();
                    _logger.LogInformation("[Bot] Disposed CallHandler for graph call {graphCallId}.", call.Id);
                }
            }
        }

        /// <summary>
        /// The get handler or throw.
        /// </summary>
        /// <param name="callLegId">
        /// The call leg id.
        /// </param>
        /// <returns>
        /// The <see cref="CallHandler"/>.
        /// </returns>
        /// <exception cref="ObjectNotFoundException">
        /// Throws an exception if handler is not found.
        /// </exception>
        private CallHandler GetHandlerOrThrow(string callLegId)
        {
            if (!CallHandlers.TryGetValue(callLegId, out CallHandler handler))
            {
                throw new ObjectNotFoundException($"call ({callLegId}) not found");
            }

            return handler;
        }

        private CallHandler ResolveCallHandlerOrThrow(string callId)
        {
            if (string.IsNullOrWhiteSpace(callId))
            {
                throw new ArgumentException("Call id is required to resolve the active media session.", nameof(callId));
            }

            if (CallHandlers.TryGetValue(callId, out var callHandler))
            {
                return callHandler;
            }

            callHandler = CallHandlers.Values.FirstOrDefault(handler =>
                string.Equals(handler.Call?.ScenarioId.ToString(), callId, StringComparison.OrdinalIgnoreCase)
                || string.Equals(handler.Call?.Id, callId, StringComparison.OrdinalIgnoreCase));

            if (callHandler != null)
            {
                return callHandler;
            }

            throw new EntityNotFoundException($"No active media call handler exists for call {callId}. The call may be persisted as established, but the Graph Communications SDK did not materialize a live call session for media operations.");
        }

        private string FormatResponseHeaders(IEnumerable<KeyValuePair<string, IEnumerable<string>>> responseHeaders)
        {
            if (responseHeaders == null)
            {
                return string.Empty;
            }

            return string.Join("; ", responseHeaders.Select(header => $"{header.Key}={string.Join(",", header.Value)}"));
        }

        private HttpRequestMessage CloneRequestMessage(HttpRequestMessage request, string payload)
        {
            if (request == null)
            {
                return null;
            }

            var clone = new HttpRequestMessage(request.Method, request.RequestUri);

            foreach (var header in request.Headers)
            {
                clone.Headers.TryAddWithoutValidation(header.Key, header.Value);
            }

            if (request.Content != null)
            {
                var mediaType = request.Content.Headers?.ContentType?.MediaType ?? "application/json";
                var content = new StringContent(payload ?? string.Empty, Encoding.UTF8, mediaType);

                foreach (var header in request.Content.Headers)
                {
                    if (string.Equals(header.Key, "Content-Type", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    content.Headers.TryAddWithoutValidation(header.Key, header.Value);
                }

                clone.Content = content;
            }

            return clone;
        }

        private async Task<bool> ReconcileNotificationPayloadAsync(string payload)
        {
            var notifications = CallingNotificationPayloadParser.Parse(payload);
            if (notifications == null || notifications.Count == 0)
            {
                var payloadPreview = string.Empty;
                if (!string.IsNullOrWhiteSpace(payload))
                {
                    payloadPreview = payload;
                    if (payload.Length > 512)
                    {
                        payloadPreview = payload.Substring(0, 512);
                    }
                }

                Log.Warning("[Bot] CallingNotificationPayloadParser returned 0 notifications. PayloadPreview={payloadPreview}", payloadPreview);
                return false;
            }

            foreach (var notification in notifications)
            {
                var state = GetNotificationState(notification.ResourceData);
                Log.Information(
                    "[Bot] Parsed notification Resource={resource} GraphCallId={graphCallId} IsParticipants={isParticipants} State={state} ResourceDataType={resourceDataType}",
                    notification.Resource,
                    notification.GraphCallId,
                    notification.IsParticipantsNotification,
                    state,
                    notification.ResourceData?.Type.ToString());
            }

            var handledAnyNotification = false;
            foreach (var notification in notifications)
            {
                try
                {
                    handledAnyNotification |= await ReconcileNotificationAsync(notification).ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    var state = GetNotificationState(notification.ResourceData);
                    _logger.LogError(
                        ex,
                        "[Bot] Raw webhook reconciliation failed for Resource={resource} GraphCallId={graphCallId} IsParticipants={isParticipants} State={state} ResourceDataType={resourceDataType}.",
                        notification.Resource,
                        notification.GraphCallId,
                        notification.IsParticipantsNotification,
                        state,
                        notification.ResourceData?.Type.ToString());

                    throw;
                }
            }

            return handledAnyNotification;
        }

        private async Task<bool> ReconcileNotificationAsync(ParsedCallingNotification notification)
        {
            var resourceData = notification.ResourceData;
            if (resourceData == null)
            {
                return false;
            }

            if (notification.IsParticipantsNotification && resourceData is JArray participants)
            {
                return await ReconcileParticipantsAsync(notification.GraphCallId, participants).ConfigureAwait(false);
            }

            var state = GetNotificationState(resourceData);
            if (string.Equals(state, Microsoft.Graph.CallState.Establishing.ToString(), StringComparison.OrdinalIgnoreCase))
            {
                return await ReconcileEstablishingCallAsync(notification.GraphCallId).ConfigureAwait(false);
            }

            if (string.Equals(state, Microsoft.Graph.CallState.Established.ToString(), StringComparison.OrdinalIgnoreCase))
            {
                return await ReconcileEstablishedCallAsync(notification.GraphCallId).ConfigureAwait(false);
            }

            if (string.Equals(state, Microsoft.Graph.CallState.Terminated.ToString(), StringComparison.OrdinalIgnoreCase))
            {
                return await ReconcileTerminatedCallAsync(notification.GraphCallId).ConfigureAwait(false);
            }

            Log.Information(
                "[Bot] Ignoring parsed notification Resource={resource} GraphCallId={graphCallId} State={state} IsParticipants={isParticipants}",
                notification.Resource,
                notification.GraphCallId,
                state,
                notification.IsParticipantsNotification);

            return false;
        }

        private async Task<bool> ReconcileEstablishingCallAsync(string graphCallId)
        {
            var scenarioCallId = await ResolveScenarioCallIdAsync(graphCallId).ConfigureAwait(false);
            if (string.IsNullOrWhiteSpace(scenarioCallId))
            {
                _logger.LogInformation("[Bot] Skipping establishing fallback for graph call {graphCallId}; no scenario call id was found.", graphCallId);
                return false;
            }

            using var scope = _serviceScopeFactory.CreateScope();
            var callRepository = scope.ServiceProvider.GetRequiredService<ICallRepository>();
            var call = await callRepository.GetItemAsync(scenarioCallId).ConfigureAwait(false);
            if (call == null)
            {
                _logger.LogInformation("[Bot] Skipping establishing fallback for graph call {graphCallId}; scenario call {callId} was not found.", graphCallId, scenarioCallId);
                return false;
            }

            if (!string.Equals(call.GraphId, graphCallId, StringComparison.OrdinalIgnoreCase))
            {
                call.GraphId = graphCallId;
                await callRepository.UpdateItemAsync(call.Id, call).ConfigureAwait(false);
                _logger.LogInformation("[Bot] Reconciled establishing webhook for graph call {graphCallId} and scenario call {callId}.", graphCallId, scenarioCallId);
            }

            return true;
        }

        private async Task<bool> ReconcileEstablishedCallAsync(string graphCallId)
        {
            _logger.LogInformation("[Bot] Attempting established fallback for graph call {graphCallId}.", graphCallId);

            var scenarioCallId = await ResolveScenarioCallIdAsync(graphCallId).ConfigureAwait(false);
            if (string.IsNullOrWhiteSpace(scenarioCallId))
            {
                _logger.LogInformation("[Bot] Skipping established fallback for graph call {graphCallId}; no scenario call id was found.", graphCallId);
                return false;
            }

            _logger.LogInformation("[Bot] Established fallback resolved scenario call {callId} for graph call {graphCallId}.", scenarioCallId, graphCallId);
            await _mediatorService.SetCallAsEstablishedAsync(scenarioCallId, graphCallId).ConfigureAwait(false);
            _logger.LogInformation("[Bot] Reconciled established webhook for graph call {graphCallId} and scenario call {callId}.", graphCallId, scenarioCallId);

            if (!CallHandlers.ContainsKey(graphCallId))
            {
                _logger.LogWarning("[Bot] Established reconciliation completed for graph call {graphCallId} without an in-memory CallHandler. This indicates CallsOnUpdated did not add the call for this notification sequence.", graphCallId);
            }

            return true;
        }

        private async Task<bool> ReconcileTerminatedCallAsync(string graphCallId)
        {
            var scenarioCallId = await ResolveScenarioCallIdAsync(graphCallId).ConfigureAwait(false);
            if (string.IsNullOrWhiteSpace(scenarioCallId))
            {
                _logger.LogInformation("[Bot] Skipping terminated fallback for graph call {graphCallId}; no scenario call id was found.", graphCallId);
                return false;
            }

            await _mediatorService.SetCallAsTerminatedAsync(scenarioCallId).ConfigureAwait(false);
            if (string.Equals(_activeScenarioCallId, scenarioCallId, StringComparison.OrdinalIgnoreCase))
            {
                _activeScenarioCallId = null;
            }

            _logger.LogInformation("[Bot] Reconciled terminated webhook for graph call {graphCallId} and scenario call {callId}.", graphCallId, scenarioCallId);
            return true;
        }

        private async Task<bool> ReconcileParticipantsAsync(string graphCallId, JArray participants)
        {
            var scenarioCallId = await ResolveScenarioCallIdAsync(graphCallId).ConfigureAwait(false);
            if (string.IsNullOrWhiteSpace(scenarioCallId))
            {
                _logger.LogInformation("[Bot] Skipping participants fallback for graph call {graphCallId}; no scenario call id was found.", graphCallId);
                return false;
            }

            var reconciledAnyParticipant = false;

            using var scope = _serviceScopeFactory.CreateScope();
            var participantRepository = scope.ServiceProvider.GetRequiredService<IParticipantStreamRepository>();

            foreach (var participant in participants.OfType<JObject>())
            {
                if (!TryBuildParticipantSnapshot(participant, scenarioCallId, out var snapshot))
                {
                    continue;
                }

                var specification = new ParticipantStreamGetFromCallSpecification(snapshot.CallId, snapshot.ParticipantGraphId);
                var existingParticipant = await participantRepository.GetFirstItemAsync(specification).ConfigureAwait(false);
                if (existingParticipant == null)
                {
                    snapshot.CreatedAt = DateTime.UtcNow;
                    await participantRepository.AddItemAsync(snapshot).ConfigureAwait(false);
                }
                else
                {
                    existingParticipant.AadId = snapshot.AadId;
                    existingParticipant.DisplayName = snapshot.DisplayName;
                    existingParticipant.PhotoUrl = snapshot.PhotoUrl;
                    existingParticipant.AudioMuted = snapshot.AudioMuted;
                    existingParticipant.IsSharingAudio = snapshot.IsSharingAudio;
                    existingParticipant.IsSharingVideo = snapshot.IsSharingVideo;
                    existingParticipant.IsSharingScreen = snapshot.IsSharingScreen;
                    existingParticipant.LeftAt = null;

                    await participantRepository.UpdateItemAsync(existingParticipant.Id, existingParticipant).ConfigureAwait(false);
                }

                reconciledAnyParticipant = true;
            }

            if (reconciledAnyParticipant)
            {
                _logger.LogInformation("[Bot] Reconciled participant webhook for graph call {graphCallId} and scenario call {callId}.", graphCallId, scenarioCallId);
            }

            return reconciledAnyParticipant;
        }

        private async Task<string> ResolveScenarioCallIdAsync(string graphCallId)
        {
            // 1. SDK in-memory call handlers (populated when SDK processes AddAsync response)
            if (CallHandlers.TryGetValue(graphCallId, out var callHandler))
            {
                _logger.LogInformation("[Bot] ResolveScenarioCallIdAsync resolved graph call {graphCallId} from in-memory CallHandlers with scenario call {callId}.", graphCallId, callHandler.Call.ScenarioId);
                return callHandler.Call.ScenarioId.ToString();
            }

            if (!string.IsNullOrWhiteSpace(_activeScenarioCallId))
            {
                _logger.LogInformation("[Bot] ResolveScenarioCallIdAsync resolved graph call {graphCallId} from in-memory active scenario call {callId}.", graphCallId, _activeScenarioCallId);
                return _activeScenarioCallId;
            }

            using var scope = _serviceScopeFactory.CreateScope();
            var callRepository = scope.ServiceProvider.GetRequiredService<ICallRepository>();

            // 2. CosmosDB lookup by GraphId (set after first established reconciliation)
            var specification = new CallGetByGraphIdSpecification(graphCallId);
            var call = await callRepository.GetFirstItemAsync(specification).ConfigureAwait(false);
            if (call != null)
            {
                _logger.LogInformation("[Bot] ResolveScenarioCallIdAsync resolved graph call {graphCallId} from persisted GraphId with scenario call {callId}.", graphCallId, call.Id);
                return call.Id;
            }

            // 3. Fallback for when SDK threw during AddAsync and GraphId hasn't been stored yet:
            //    look up the service's current active call via the registered service ID.
            if (!string.IsNullOrWhiteSpace(Id))
            {
                var serviceRepository = scope.ServiceProvider.GetRequiredService<IServiceRepository>();
                var service = await serviceRepository.GetItemAsync(Id).ConfigureAwait(false);
                if (!string.IsNullOrWhiteSpace(service?.CallId))
                {
                    _logger.LogInformation("[Bot] ResolveScenarioCallIdAsync resolved graph call {graphCallId} from active service {serviceId} with scenario call {callId}.", graphCallId, Id, service.CallId);
                    return service.CallId;
                }
            }

            _logger.LogWarning("[Bot] ResolveScenarioCallIdAsync could not resolve a scenario call id for graph call {graphCallId}.", graphCallId);
            return null;
        }

        private string GetNotificationState(JToken resourceData)
        {
            if (!(resourceData is JObject resourceObject))
            {
                return null;
            }

            return resourceObject.Value<string>("state");
        }

        private bool TryBuildParticipantSnapshot(JObject participant, string callId, out ParticipantStream snapshot)
        {
            snapshot = null;

            var identity = participant["info"]?["identity"] as JObject;
            if (identity == null)
            {
                return false;
            }

            var user = identity["user"] as JObject;
            var guest = identity["guest"] as JObject;
            var application = identity["application"] as JObject;

            if (application != null && string.Equals(application.Value<string>("id"), _config.AadAppId, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            var isLiveEventBot = application != null
                && string.Equals(application.Value<string>("displayName"), "live", StringComparison.OrdinalIgnoreCase);

            if (user == null && guest == null && !isLiveEventBot)
            {
                return false;
            }

            var participantGraphId = participant.Value<string>("id") ?? participant["info"]?.Value<string>("participantId");
            if (string.IsNullOrWhiteSpace(participantGraphId))
            {
                return false;
            }

            var principal = user ?? guest ?? application;
            var aadId = principal?.Value<string>("id");
            var displayName = principal?.Value<string>("displayName");
            var mediaStreams = participant["mediaStreams"] as JArray;

            snapshot = new ParticipantStream
            {
                AadId = aadId,
                CallId = callId,
                ParticipantGraphId = participantGraphId,
                DisplayName = displayName,
                PhotoUrl = string.IsNullOrWhiteSpace(aadId)
                    ? null
                    : $"https://{_config.MainApiUrl}/api/participant/photo/{aadId}",
                Type = isLiveEventBot ? ResourceType.LiveEvent : ResourceType.Participant,
                State = StreamState.Disconnected,
                IsHealthy = true,
                HealthMessage = string.Empty,
                AudioMuted = participant.Value<bool?>("isMuted") ?? false,
                IsSharingAudio = HasSendCapableMediaStream(mediaStreams, Modality.Audio.ToString()),
                IsSharingVideo = HasSendCapableMediaStream(mediaStreams, Modality.Video.ToString()),
                IsSharingScreen = HasSendCapableMediaStream(mediaStreams, Modality.VideoBasedScreenSharing.ToString()),
                Details = new ParticipantStreamDetails(),
            };

            return true;
        }

        private void LogGraphModelResolution(string typeName)
        {
            var matches = AppDomain.CurrentDomain
                .GetAssemblies()
                .Select(assembly => new
                {
                    Assembly = assembly,
                    Type = assembly.GetType(typeName, throwOnError: false, ignoreCase: false),
                })
                .Where(entry => entry.Type != null)
                .ToList();

            if (!matches.Any())
            {
                _logger.LogInformation("[Bot] Graph model resolution check: {typeName} was not found in loaded assemblies.", typeName);
                return;
            }

            foreach (var match in matches)
            {
                var constructors = string.Join(", ", match.Type.GetConstructors().Select(constructor => constructor.ToString()));

                _logger.LogInformation(
                    "[Bot] Graph model resolution check: {typeName} resolved from assembly {assemblyName}. Constructors={constructors}",
                    typeName,
                    match.Assembly.FullName,
                    string.IsNullOrWhiteSpace(constructors) ? "<none>" : constructors);
            }
        }

        private bool HasSendCapableMediaStream(JArray mediaStreams, string modality)
        {
            if (mediaStreams == null)
            {
                return false;
            }

            return mediaStreams
                .OfType<JObject>()
                .Any(stream =>
                    string.Equals(stream.Value<string>("mediaType"), modality, StringComparison.OrdinalIgnoreCase)
                    && IsSendCapableDirection(stream.Value<string>("direction")));
        }

        private bool IsSendCapableDirection(string direction)
        {
            return string.Equals(direction, MediaDirection.SendOnly.ToString(), StringComparison.OrdinalIgnoreCase)
                || string.Equals(direction, MediaDirection.SendReceive.ToString(), StringComparison.OrdinalIgnoreCase);
        }
        #endregion
    }
}
