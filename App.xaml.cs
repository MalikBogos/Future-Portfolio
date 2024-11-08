using FuturePortfolio.Core;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using System.Configuration;
using System.Data;
using System.Windows;
using Application = System.Windows.Application;
using FuturePortfolio.Core;
using Microsoft.EntityFrameworkCore;

namespace FuturePortfolio
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public static IHost? Host { get; private set; }

        public App()
        {
            Host = CreateHostBuilder().Build();
            InitializeDatabase();
        }

        private void InitializeDatabase()
        {
            using var scope = Host?.Services.CreateScope();
            var db = scope?.ServiceProvider.GetRequiredService<FuturePortfolioDbContext>();

            if (db != null)
            {
                // This will create the database if it doesn't exist
                // and apply any pending migrations
                db.Database.Migrate();
            }
        }

        public static IHostBuilder CreateHostBuilder(string[]? args = null) =>
            Microsoft.Extensions.Hosting.Host.CreateDefaultBuilder(args)
                .ConfigureServices((hostContext, services) =>
                {
                    services.AddDbContextPool<FuturePortfolioDbContext>(options =>
                        options.UseSqlServer(
                            "Server=localhost;Database=FuturePortfolio;Integrated Security=True;TrustServerCertificate=True;"));

                    services.AddScoped<ISpreadsheetService, SpreadsheetService>();
                    services.AddScoped<IFileOperationsService, FileOperationsService>();
                    services.AddTransient<MainViewModel>();
                });

        protected override async void OnStartup(StartupEventArgs e)
        {
            await Host!.StartAsync();
            base.OnStartup(e);
        }

        protected override async void OnExit(ExitEventArgs e)
        {
            if (Host != null)
            {
                await Host.StopAsync();
                Host.Dispose();
            }
            base.OnExit(e);
        }
    }
}