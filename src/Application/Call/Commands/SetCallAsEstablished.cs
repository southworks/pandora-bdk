// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Application.Interfaces.Persistance;
using Application.Participants.Specifications;
using Domain.Entities;
using Domain.Entities.Parts;
using Domain.Enums;
using Domain.Exceptions;
using FluentValidation;
using MediatR;
using Microsoft.Extensions.Logging;
using static Domain.Constants.Constants;

namespace Application.Call.Commands
{
    public class SetCallAsEstablished
    {
        public class SetCallAsEstablishedCommand : IRequest<SetCallAsEstablishedCommandResponse>
        {
            public string CallId { get; set; }

            public string GraphCallId { get; set; }
        }

        public class SetCallAsEstablishedCommandResponse
        {
            public string Id { get; set; }
        }

        public class SetCallAsEstablishedCommandValidator : AbstractValidator<SetCallAsEstablishedCommand>
        {
            public SetCallAsEstablishedCommandValidator()
            {
                RuleFor(x => x.CallId)
                    .NotEmpty();
                RuleFor(x => x.GraphCallId)
                    .NotEmpty();
            }
        }

        public class SetCallAsEstablishedCommandHandler : IRequestHandler<SetCallAsEstablishedCommand, SetCallAsEstablishedCommandResponse>
        {
            private readonly ICallRepository _callRepository;
            private readonly IParticipantStreamRepository _participantStreamRepository;
            private readonly ILogger<SetCallAsEstablishedCommandHandler> _logger;

            public SetCallAsEstablishedCommandHandler(
                ICallRepository callRepository,
                IParticipantStreamRepository participantStreamRepository,
                ILogger<SetCallAsEstablishedCommandHandler> logger)
            {
                _callRepository = callRepository;
                _participantStreamRepository = participantStreamRepository;
                _logger = logger;
            }

            public async Task<SetCallAsEstablishedCommandResponse> Handle(SetCallAsEstablishedCommand request, CancellationToken cancellationToken)
            {
                var response = new SetCallAsEstablishedCommandResponse();

                var entity = await _callRepository.GetItemAsync(request.CallId);
                if (entity == null)
                {
                    _logger.LogError("Call with id {id} was not found", request.CallId);
                    throw new EntityNotFoundException($"Call with id  {request.CallId} was not found");
                }

                var wasEstablished = entity.State == CallState.Established;

                entity.State = CallState.Established;
                if (!wasEstablished)
                {
                    entity.StartedAt = DateTime.UtcNow;
                }

                entity.GraphId = request.GraphCallId;

                await _callRepository.UpdateItemAsync(entity.Id, entity);
                await EnsureDefaultParticipantsStreams(request.CallId);

                response.Id = entity.Id;

                return response;
            }

            private async Task EnsureDefaultParticipantsStreams(string callId)
            {
                var specification = new ParticipantsStreamsGetFromCallSpecification(callId, archived: false);
                var existingStreams = (await _participantStreamRepository.GetItemsAsync(specification)).ToList();
                var insertTasks = new List<Task>();

                foreach (var participant in CreateDefaultParticipantStreams())
                {
                    var alreadyExists = existingStreams.Any(existing =>
                        existing.Type == participant.Type
                        && string.Equals(existing.DisplayName, participant.DisplayName, StringComparison.Ordinal));

                    if (alreadyExists)
                    {
                        continue;
                    }

                    participant.CallId = callId;
                    insertTasks.Add(_participantStreamRepository.AddItemAsync(participant));
                }

                await Task.WhenAll(insertTasks);
            }

            private static IEnumerable<ParticipantStream> CreateDefaultParticipantStreams()
            {
                return new[]
                {
                    new ParticipantStream
                    {
                        ParticipantGraphId = Guid.NewGuid().ToString(),
                        DisplayName = DefaultParticipantsDisplayNames.PrimarySpeaker,
                        Type = ResourceType.PrimarySpeaker,
                        State = StreamState.Disconnected,
                        IsHealthy = true,
                        Details = new ParticipantStreamDetails(),
                    },
                    new ParticipantStream
                    {
                        ParticipantGraphId = Guid.NewGuid().ToString(),
                        DisplayName = DefaultParticipantsDisplayNames.ScreenShare,
                        Type = ResourceType.Vbss,
                        State = StreamState.Disconnected,
                        IsHealthy = true,
                        Details = new ParticipantStreamDetails(),
                    },
                };
            }
        }
    }
}
