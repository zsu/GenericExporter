using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.Text;

namespace GenericExporter
{
    public static class ServiceCollectionExtensions
    {
        public static IServiceCollection AddGenericExporter(this IServiceCollection services)
        {
            Register(services);
            return services;
        }

        private static void Register(IServiceCollection services)
        {
            services.AddTransient<IExporter, Exporter>();
        }
    }
}
