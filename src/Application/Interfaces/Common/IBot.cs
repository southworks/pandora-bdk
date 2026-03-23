// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
using System.Net.Http;
using System.Threading.Tasks;
using Application.Common.Models;
using Application.Common.Models.Api;
using static Application.Service.Commands.DoInviteBot;

namespace Application.Interfaces.Common
{
    public interface IBot
    {
        string Id { get;  }

        string VirtualMachineName { get; set; }

        Task InviteBotAsync(DoInviteBotCommand command);

        Task ProcessNotificationAsync(HttpRequestMessage request);

        Task EndCallAsync(string callGraphId);

        void StartInjection(StartStreamInjectionBody startStreamInjectionBody);

        void StopInjection(string callId);

        void SetInjectionVolume(string callId, SetInjectionVolumeRequest injectionVolumeRequest);

        void DisplayInjection(string callId);

        void HideInjection(string callId);

        StartStreamExtractionResponse StartExtraction(StartStreamExtractionBody streamBody);

        void StopExtraction(StopStreamExtractionBody streamBody);

        Task MuteBotAsync(string callId);

        Task UnmuteBotAsync(string callId);

        Task RegisterServiceAsync(string virtualMachineName);

        Task UnregisterServiceAsync(string virtualMachineName);
    }
}