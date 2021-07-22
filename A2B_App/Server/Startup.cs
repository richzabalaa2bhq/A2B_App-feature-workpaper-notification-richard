using A2B_App.Server.Data;
using A2B_App.Server.Models;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.AspNetCore.Mvc.NewtonsoftJson;
using Microsoft.AspNetCore.Identity;
using System.Linq;
using Microsoft.Extensions.FileProviders;
using Microsoft.AspNetCore.Http;
using System.IO;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.AspNetCore.ApiAuthorization.IdentityServer;
using Microsoft.IdentityModel.Tokens;
using Microsoft.AspNetCore.Components.Routing;
using Quartz;
using System.Collections.Specialized;
using Quartz.Impl;
using System.Xml.Schema;
using System.Text;
using System;
using A2B_App.Server.JobScheduler;
using Quartz.Spi;

namespace A2B_App.Server
{
    public class Startup
    {
        public Startup(IConfiguration configuration)
        {
            Configuration = configuration;
        }

        public IConfiguration Configuration { get; }

        // This method gets called by the runtime. Use this method to add services to the container.
        // For more information on how to configure your application, visit https://go.microsoft.com/fwlink/?LinkID=398940
        public void ConfigureServices(IServiceCollection services)
        {

            services.AddCors(options =>
            {
                options.AddPolicy("CorsPolicy",
                    builder => builder
                    //.WithOrigins()
                    //.AllowAnyOrigin()
                    .AllowAnyMethod()
                    .AllowAnyHeader()
                    .SetIsOriginAllowed(origin => true)
                    //.WithMethods("GET, PATCH, DELETE, PUT, POST, OPTIONS")
                    );
            });

            services.AddDbContext<ApplicationDbContext>(options => options.UseMySQL(Configuration.GetConnectionString("IdentityConnection")));;
            services.AddDbContext<SoxContext>(options => options.UseMySQL(Configuration.GetConnectionString("SampleSelectionCon")));
            services.AddDbContext<SmsContext>(options => options.UseMySQL(Configuration.GetConnectionString("SmsCon")));
            services.AddDbContext<VerificationContext>(options => options.UseMySQL(Configuration.GetConnectionString("VerificationCon")));
            services.AddDbContext<TimeContext>(options => options.UseMySQL(Configuration.GetConnectionString("TimeCon")));
            services.AddDbContext<MeetingContext>(options => options.UseMySQL(Configuration.GetConnectionString("MeetingCon")));
            services.AddDbContext<UserContext>(options => options.UseMySQL(Configuration.GetConnectionString("UserCon")));
            services.AddDbContext<NotificationContext>(options => options.UseMySQL(Configuration.GetConnectionString("NotificationCon")));

            //Add swagger
            //------------------------------------------------
            services.AddSwaggerGen(c =>
            {
                c.SwaggerDoc("v1", new Microsoft.OpenApi.Models.OpenApiInfo { Title = "A2BHQ API", Version = "v1" });
            });
            //------------------------------------------------

            services.AddDefaultIdentity<ApplicationUser>(options => options.SignIn.RequireConfirmedAccount = true)                
                .AddRoles<IdentityRole>()
                .AddEntityFrameworkStores<ApplicationDbContext>();

            // Configure identity server to put the role claim into the id token 
            // and the access token and prevent the default mapping for roles 
            // in the JwtSecurityTokenHandler.
            services.AddIdentityServer(options =>
            {
                options.IssuerUri = Configuration["IdentityServer:IssuerUri"];
            })
            //.AddDeveloperSigningCredential()
            //.AddInMemoryApiResources(Data.ResourceManager.GetApiResources())
            //.AddInMemoryClients(Data.ClientManager.Clients)
            //.AddInMemoryIdentityResources(Data.ResourceManager.GetIdentityResources())
            .AddApiAuthorization<ApplicationUser, ApplicationDbContext>(options =>
            {
                //options.Clients.AddRange(Configuration.GetSection("IdentityServer:Clients").Get<IdentityServer4.Models.Client[]>());
                options.IdentityResources["openid"].UserClaims.Add("role");
                options.ApiResources.Single().UserClaims.Add("role");
            });
           

            // Need to do this as it maps "role" to ClaimTypes.Role and causes issues
            System.IdentityModel.Tokens.Jwt.JwtSecurityTokenHandler.DefaultInboundClaimTypeMap.Remove("role");

            services.AddAuthentication()
                //.AddJwtBearer(options =>
                //{
                //    options.TokenValidationParameters = new TokenValidationParameters
                //    {
                //        ValidateIssuer = true,
                //        ValidateAudience = true,
                //        ValidateIssuerSigningKey = true,
                //        ValidIssuer = Configuration["TokenSettings:Issuer"],
                //        ValidAudience = Configuration["TokenSettings:Audience"],
                //        ValidateLifetime = true,
                //        LifetimeValidator = LifetimeValidator,
                //        IssuerSigningKey = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(Configuration["TokenSettings:IssuerSigningKey"]))
                //    };
                //})     
                //.AddOpenIdConnect("oidc", options =>
                //{
                //    options.SignInScheme = "Cookies";

                //    options.Authority = Configuration["IdentityServer:IssuerUri"];
                //    options.RequireHttpsMetadata = false;

                //    options.ClientId = "A2B_App.Client";
                //    options.ClientSecret = "eb300de4-add9-42f4-a3ac-abd3c60f1919";

                //    options.SaveTokens = true;
                //    options.GetClaimsFromUserInfoEndpoint = true;

                //    options.Scope.Add("api1");
                //    options.Scope.Add("offline_access");
                //})
                .AddIdentityServerJwt();

            

            //services.AddAuthentication(options =>
            //{
            //    options.DefaultAuthenticateScheme = JwtBearerDefaults.AuthenticationScheme;
            //})
            //.AddJwtBearer(options =>
            //{
            //    options.TokenValidationParameters = new TokenValidationParameters
            //    {
            //        ValidateIssuer = true,
            //        ValidateAudience = true,
            //        ValidateIssuerSigningKey = true,
            //        ValidIssuer = Configuration["TokenSettings:Issuer"],
            //        ValidAudience = Configuration["TokenSettings:Audience"],
            //        ValidateLifetime = true,
            //        LifetimeValidator = LifetimeValidator,
            //        IssuerSigningKey = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(Configuration["TokenSettings:IssuerSigningKey"]))
            //    };
            //})
            //.AddIdentityServerJwt();


            //services.Configure<JwtBearerOptions>(
            //    IdentityServerJwtConstants.IdentityServerJwtBearerScheme,
            //    options =>
            //    {
            //        options.RequireHttpsMetadata = true;
            //        options.SaveToken = true;
            //        options.TokenValidationParameters = new TokenValidationParameters
            //        {
            //            ValidateIssuer = true,
            //            ValidateAudience = true,
            //            ValidateLifetime = true,
            //            ValidateIssuerSigningKey = true,
            //            ValidIssuer = "https://ddohq.com/",
            //            ValidAudience = "https://ddohq.com/",
            //        };
            //    });

            services.AddControllersWithViews().AddNewtonsoftJson(options =>
                options.SerializerSettings.ReferenceLoopHandling = Newtonsoft.Json.ReferenceLoopHandling.Ignore
             );



            services.AddRazorPages();

            //Job Scheduler
            services.AddSingleton(provider => GetScheduler());
            // Add Quartz services
            services.AddHostedService<QuartzHostedService>();
            services.AddSingleton<IJobFactory, SingletonJobFactory>();
            services.AddSingleton<ISchedulerFactory, StdSchedulerFactory>();

            services.AddSingleton<SodSoxRoxJob>();
            services.AddSingleton<MeetingNotificationJobs>();
            services.AddSingleton<Meeting3PMNotificationJobs>();
            services.AddSingleton<BizDevNotificationJobs>();
            services.AddSingleton(new JobSchedule(
                jobType: typeof(MeetingNotificationJobs),
                cronExpression: "0 0/5 * * * ?")); // run every 5 min
            services.AddSingleton(new JobSchedule(
                jobType: typeof(BizDevNotificationJobs),
                cronExpression: "0 0/5 * * * ?")); // run every 5 min
            services.AddSingleton(new JobSchedule(
                jobType: typeof(Meeting3PMNotificationJobs),
                  cronExpression: "0 0 15 * * ?")); // run every 3:00 PM
                //  cronExpression: "0 0/5 * * * ?")); // run every 5 min
        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
        {
            app.UseCors("CorsPolicy");

            if (env.IsDevelopment())
            {
                //------------------------------------------------
                //enable for API Testing
                //------------------------------------------------
                app.UseSwagger();
                app.UseSwaggerUI(c =>
                {
                    c.SwaggerEndpoint("/swagger/v1/swagger.json", "A2BHQ API");
                });
                //------------------------------------------------

                app.UseDeveloperExceptionPage();
                app.UseDatabaseErrorPage();
                app.UseWebAssemblyDebugging();
            }
            else
            {
                app.UseExceptionHandler("/Error");
                // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
                app.UseHsts();
            }

            app.UseHttpsRedirection();
            app.UseBlazorFrameworkFiles();
            app.UseStaticFiles();
            app.UseStaticFiles(new StaticFileOptions()
            {
                FileProvider = new PhysicalFileProvider(Path.Combine(Directory.GetCurrentDirectory(), @"include/upload/image")),
                RequestPath = new PathString("/include/upload/image")
            });
            app.UseRouting();
            
            app.UseIdentityServer();
            app.UseAuthentication();
            app.UseAuthorization();

            app.UseEndpoints(endpoints =>
            {
                //endpoints.MapDefaultControllerRoute();
                endpoints.MapRazorPages();
                endpoints.MapControllers();
                endpoints.MapFallbackToFile("index.html");
            });


           

        }

        private IScheduler GetScheduler()
        {
            var properties = new NameValueCollection
            {
                ["quartz.scheduler.instanceName"] = "A2B Job Scheduler",
                ["quartz.jobStore.type"] = "Quartz.Simpl.RAMJobStore, Quartz",
                ["quartz.threadPool.maxConcurrency"] = "3"
            };
            var schedulerFactory = new StdSchedulerFactory();
            var scheduler = schedulerFactory.GetScheduler().Result;
            scheduler.Start();
            return scheduler;

        }

        private bool LifetimeValidator(DateTime? notBefore, DateTime? expires, SecurityToken token, TokenValidationParameters @params)
        {
            if (expires != null)
                return expires > DateTime.UtcNow;
            return false;
        }
    



    }
}
