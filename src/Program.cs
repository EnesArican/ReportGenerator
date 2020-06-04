using Microsoft.Extensions.DependencyInjection;
using ReportGenerator.Interfaces;
using ReportGenerator.Services;

namespace ReportGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create service collection and configure our services
            var services = ConfigureServices();
            // Generate a provider
            var serviceProvider = services.BuildServiceProvider();

            // Kick off our actual code
            serviceProvider.GetService<ConsoleApp>().Run();
        }

        private static IServiceCollection ConfigureServices() 
        {
            IServiceCollection services = new ServiceCollection();
            services.AddTransient<ConsoleApp>();
            services.AddSingleton<IFileHandlerService, FileHandlerService>();
            services.AddSingleton<IDataReaderService, DataReaderService>();
            services.AddSingleton<IDataWriterService, DataWriterService>();
            services.AddSingleton<IFormatterService, FormatterService>();
            services.AddSingleton<IRangeEditorService, RangeEditorService>();
            return services;
        }
    }
}
