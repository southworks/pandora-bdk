// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using BotService.Configuration;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Server.Kestrel.Core;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Serilog;

namespace BotService.Infrastructure.Extensions
{
    public static class KestrelServerOptionsExtensions
    {
        public static void ConfigureEndpoints(this KestrelServerOptions options, X509Certificate2 certificate)
        {
            var configuration = options.ApplicationServices.GetRequiredService<IConfiguration>();
            Log.Information("[Kestrel] Resolving configured endpoints.");

            var endpoints = configuration.GetSection("HttpServer:Endpoints")
                .GetChildren()
                .ToDictionary(section => section.Key, section =>
                {
                    var endpoint = new EndpointConfiguration();
                    section.Bind(endpoint);
                    return endpoint;
                });

            foreach (var endpoint in endpoints)
            {
                var config = endpoint.Value;
                var port = config.Port ?? (config.Scheme == "https" ? 443 : 80);
                Log.Information("[Kestrel] Configuring endpoint {endpointName}. Host={host} Scheme={scheme} Port={port}", endpoint.Key, config.Host, config.Scheme, port);

                var ipAddresses = new List<IPAddress>();
                if (config.Host == "localhost")
                {
                    ipAddresses.Add(IPAddress.IPv6Loopback);
                    ipAddresses.Add(IPAddress.Loopback);
                }
                else if (IPAddress.TryParse(config.Host, out var address))
                {
                    ipAddresses.Add(address);
                }
                else
                {
                    ipAddresses.Add(IPAddress.IPv6Any);
                }

                foreach (var address in ipAddresses)
                {
                    Log.Information("[Kestrel] Binding endpoint {endpointName} to address {address}:{port}", endpoint.Key, address, port);
                    options.Listen(
                        address,
                        port,
                        listenOptions =>
                        {
                            if (config.Scheme == "https")
                            {
                                listenOptions.UseHttps(certificate);
                            }
                        });
                }
            }
        }
    }
}
