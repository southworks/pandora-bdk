// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
using System.Reflection;
using Application;
using Application.Common.Config;
using Application.Interfaces.Common;
using Application.Interfaces.Persistance;
using AutoMapper;
using BotServiceParticipantMappingProfile = BotService.Application.Participant.MappingProfile;
using BotService.Application.Core;
using BotService.Configuration;
using BotService.Infrastructure.Client;
using BotService.Infrastructure.Common;
using BotService.Infrastructure.Common.Logging;
using BotService.Infrastructure.Core;
using BotService.Infrastructure.Pipelines;
using BotService.Infrastructure.Services;
using Infrastructure.Core.Common;
using Infrastructure.Core.CosmosDbData.Extensions;
using Infrastructure.Core.CosmosDbData.Repository;
using Infrastructure.Core.Services;
using Microsoft.ApplicationInsights.Extensibility;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Authorization;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Graph.Communications.Client;
using Microsoft.Graph.Communications.Common.Telemetry;
using Serilog;

namespace BotService
{
    public class Startup
    {
        private readonly ILogger<Startup> _logger;

        public Startup(
            IConfiguration configuration,
            ILogger<Startup> logger)
        {
            Configuration = configuration;
            _logger = logger;
        }

        public IConfiguration Configuration { get; }

        // This method gets called by the runtime. Use this method to add services to the container.
        public void ConfigureServices(IServiceCollection services)
        {
            Log.Information("[Startup] ConfigureServices starting.");

            var environment = new HostEnvironment();
            services.AddSingleton<IHostEnvironment>(environment);
            Log.Information("[Startup] Registered host environment. IsLocal={isLocal}", environment.IsLocal());

            var appConfiguration = Configuration.GetSection("Settings").Get<AppConfiguration>();
            services.AddSingleton<IAppConfiguration>(appConfiguration);
            Log.Information("[Startup] Loaded app configuration. ServiceDnsName={serviceDnsName} CosmosEndpointConfigured={hasCosmosEndpoint}", appConfiguration?.BotConfiguration?.ServiceDnsName, !string.IsNullOrWhiteSpace(appConfiguration?.CosmosDbConfiguration?.EndpointUrl));

            var graphLogger = new GraphLogger(typeof(Program).Assembly.GetName().Name, redirectToTrace: true);
            services.AddSingleton<IGraphLogger>(graphLogger);
            Log.Information("[Startup] Registered Graph logger.");

            var sampleObserver = new SampleObserver(graphLogger);
            services.AddSingleton(sampleObserver);
            services.AddSingleton<PipelineBusObserver>();
            Log.Information("[Startup] Registered pipeline observers.");

            services.AddSingleton<ICommunicationsClient>(x =>
            {
                var appConfiguration = x.GetService<IAppConfiguration>();
                var graphLogger = x.GetService<IGraphLogger>();
                var loggerFactory = x.GetService<ILoggerFactory>();

                var clientBuilder = new GraphCommunicationsClientBuilder(
                    appConfiguration,
                    graphLogger,
                    loggerFactory.CreateLogger<GraphCommunicationsClientBuilder>());

                return clientBuilder.Build();
            });
            Log.Information("[Startup] Registered communications client factory.");

            if (!environment.IsLocal())
            {
                services.AddAuthentication(JwtBearerDefaults.AuthenticationScheme).AddJwtBearer(options =>
                {
                    options.Authority = $"{appConfiguration.AzureAdConfiguration.Instance}{appConfiguration.AzureAdConfiguration.TenantId}/v2.0";
                    options.Audience = $"{appConfiguration.BotServiceAuthenticationConfiguration.BotServiceApiClientId}";
                });

                Log.Information("[Startup] Configured JWT bearer authentication.");
            }

            services.AddApplication();
            Log.Information("[Startup] Registered application services.");

            Log.Information("[Startup] Registering BotService AutoMapper profile explicitly.");
            services.AddAutoMapper(config => config.AddProfile<BotServiceParticipantMappingProfile>());
            Log.Information("[Startup] Registered AutoMapper.");

            // register CosmosDB client and data repositories
            services.AddCosmosDb(
                appConfiguration.CosmosDbConfiguration.EndpointUrl,
                appConfiguration.CosmosDbConfiguration.PrimaryKey,
                appConfiguration.CosmosDbConfiguration.DatabaseName);
            Log.Information("[Startup] Registered Cosmos DB services. DatabaseName={databaseName}", appConfiguration?.CosmosDbConfiguration?.DatabaseName);

            services.AddScoped<IServiceRepository, ServiceRepository>();
            services.AddScoped<ICallRepository, CallRepository>();
            services.AddScoped<IParticipantStreamRepository, ParticipantStreamRepository>();
            services.AddScoped<IStreamRepository, StreamRepository>();
            services.AddScoped<IAzStorageHandler, AzStorageHandler>();
            Log.Information("[Startup] Registered repositories and storage handler.");

            services.AddSingleton<IMediatorService, MediatorService>();
            services.AddSingleton<IMediaProcessorFactory, GStreamerMediaProcessorFactory>();
            services.AddSingleton<GstreamerClockProvider>();
            services.AddSingleton<IMediaHandlerFactory, MediaHandlerFactory>();
            services.AddSingleton<IBot, Bot>();
            Log.Information("[Startup] Registered bot and media services.");

            services.AddApplicationInsightsTelemetry();

            services.AddSingleton<ITelemetryInitializer, CloudRoleNameTelemetryInitializer>();
            Log.Information("[Startup] Registered Application Insights services.");

            services.AddScoped<IInjectionUrlHelper, InjectionUrlHelper>();
            services.AddScoped<IExtractionUrlHelper, ExtractionUrlHelper>();
            Log.Information("[Startup] Registered stream URL helpers.");

            services.AddMvc(config =>
            {
                if (!environment.IsLocal())
                {
                    var policy = new AuthorizationPolicyBuilder()
                        .RequireAuthenticatedUser()
                        .RequireRole("BotService.AccessAll")
                        .Build();

                    config.Filters.Add(new AuthorizeFilter(policy));
                }
            }).SetCompatibilityVersion(CompatibilityVersion.Version_2_1);

            Log.Information("[Startup] ConfigureServices completed.");
        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        public void Configure(
            IApplicationBuilder app,
            IHostEnvironment env)
        {
            if (env.IsProduction())
            {
                app.UseHsts();
            }
            else
            {
                app.UseDeveloperExceptionPage();
            }

            app.UseAuthentication();
            app.ConfigureExceptionHandler(_logger, env.IsProduction(), "BotServiceApi");
            app.UseMvc();
        }
    }
}
