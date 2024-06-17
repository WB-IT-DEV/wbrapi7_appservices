using Microsoft.Extensions.Configuration;
using wbrapi7_appservices.Data;
using wbrapi7_appservices.Repositories;
using Microsoft.EntityFrameworkCore;


var builder = WebApplication.CreateBuilder(args);



// Add services to the container.
//builder.Services.AddDbContext<WBRDataContext>(options => options.UseSqlServer(Configuration.GetConnectionString("DefaultConnection")));
builder.Services.AddDbContext<WBRDataContext>(options =>
    options.UseSqlServer(builder.Configuration.GetConnectionString("DefaultConnection")));

builder.Services.AddScoped<IWBRDataRepository, WBRDataRepository>();
builder.Services.AddControllers();

builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowWebApp",
        policyBuilder => policyBuilder
            .WithOrigins("http://localhost:3000") // Specify the origin of your front-end app
            .AllowAnyHeader()
            .AllowAnyMethod());
});


// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();


Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("Mgo+DSMBMAY9C3t2UFhhQlJBfV5AQmBIYVp/TGpJfl96cVxMZVVBJAtUQF1hTX5QdEZjWXxZcX1TQmBb");

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseCors("AllowWebApp");

app.UseAuthorization();

app.MapControllers();

app.Run();
