using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Application.Call.Commands;
using Application.Interfaces.Persistance;
using Ardalis.Specification;
using Domain.Constants;
using Domain.Entities;
using Domain.Entities.Base;
using Domain.Enums;
using DomainCall = Domain.Entities.Call;
using MediatR;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.DependencyInjection;
using Xunit;
using static Application.Call.Commands.SetCallAsEstablished;

namespace Application.UnitTests
{
    public class SetCallAsEstablishedIntegrationTests
    {
        [Fact]
        public async Task Handle_ShouldSetCallAsEstablished_AndCreateDefaultStreams_WithoutDuplicates()
        {
            var services = new ServiceCollection();
            services.AddSingleton(typeof(ILogger<>), typeof(NullLogger<>));
            services.AddApplication();

            var callRepository = new InMemoryCallRepository();
            var participantRepository = new InMemoryParticipantStreamRepository();

            services.AddSingleton<ICallRepository>(callRepository);
            services.AddSingleton<IParticipantStreamRepository>(participantRepository);

            var serviceProvider = services.BuildServiceProvider();
            var mediator = serviceProvider.GetRequiredService<IMediator>();

            var callId = Guid.NewGuid().ToString();
            var graphCallId = Guid.NewGuid().ToString();

            await callRepository.AddItemAsync(new DomainCall
            {
                Id = callId,
                State = CallState.Establishing,
            });

            await mediator.Send(new SetCallAsEstablishedCommand
            {
                CallId = callId,
                GraphCallId = graphCallId,
            });

            var updatedCall = await callRepository.GetItemAsync(callId);
            Assert.NotNull(updatedCall);
            Assert.Equal(CallState.Established, updatedCall.State);
            Assert.Equal(graphCallId, updatedCall.GraphId);

            var streamsAfterFirstCall = participantRepository.Items.Where(x => x.CallId == callId).ToList();
            Assert.Equal(2, streamsAfterFirstCall.Count);
            Assert.Contains(streamsAfterFirstCall, x => x.DisplayName == Constants.DefaultParticipantsDisplayNames.PrimarySpeaker && x.Type == ResourceType.PrimarySpeaker);
            Assert.Contains(streamsAfterFirstCall, x => x.DisplayName == Constants.DefaultParticipantsDisplayNames.ScreenShare && x.Type == ResourceType.Vbss);

            // Run the command twice to validate idempotent default stream creation.
            await mediator.Send(new SetCallAsEstablishedCommand
            {
                CallId = callId,
                GraphCallId = graphCallId,
            });

            var streamsAfterSecondCall = participantRepository.Items.Where(x => x.CallId == callId).ToList();
            Assert.Equal(2, streamsAfterSecondCall.Count);
        }

        private abstract class InMemoryRepositoryBase<T> : IRepository<T>
            where T : CosmosDbEntity
        {
            protected readonly List<T> Storage = new List<T>();

            public Task AddItemAsync(T item)
            {
                if (string.IsNullOrWhiteSpace(item.Id))
                {
                    item.Id = Guid.NewGuid().ToString();
                }

                Storage.Add(item);
                return Task.CompletedTask;
            }

            public Task DeleteItemAsync(string id)
            {
                var item = Storage.FirstOrDefault(x => x.Id == id);
                if (item != null)
                {
                    Storage.Remove(item);
                }

                return Task.CompletedTask;
            }

            public Task<T> GetFirstItemAsync(ISpecification<T> specification)
            {
                return Task.FromResult(ApplySpecification(specification).FirstOrDefault());
            }

            public Task<T> GetItemAsync(string id)
            {
                return Task.FromResult(Storage.FirstOrDefault(x => x.Id == id));
            }

            public Task<int> GetItemsCountAsync(ISpecification<T> specification)
            {
                return Task.FromResult(ApplySpecification(specification).Count());
            }

            public Task<IEnumerable<T>> GetItemsAsync(ISpecification<T> specification)
            {
                return Task.FromResult(ApplySpecification(specification));
            }

            public Task UpdateItemAsync(string id, T item)
            {
                var index = Storage.FindIndex(x => x.Id == id);
                if (index >= 0)
                {
                    Storage[index] = item;
                }
                else
                {
                    Storage.Add(item);
                }

                return Task.CompletedTask;
            }

            protected virtual IEnumerable<T> ApplySpecification(ISpecification<T> specification)
            {
                if (specification == null)
                {
                    return Storage;
                }

                IEnumerable<T> query = Storage;
                foreach (var whereExpression in specification.WhereExpressions)
                {
                    query = query.AsQueryable().Where(whereExpression).ToList();
                }

                return query;
            }
        }

        private sealed class InMemoryCallRepository : InMemoryRepositoryBase<DomainCall>, ICallRepository
        {
        }

        private sealed class InMemoryParticipantStreamRepository : InMemoryRepositoryBase<ParticipantStream>, IParticipantStreamRepository
        {
            public IReadOnlyList<ParticipantStream> Items => Storage;
        }
    }
}