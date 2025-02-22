using Burse.Data;
using Burse.Services.Abstractions;
using Burse.Services;

using Microsoft.EntityFrameworkCore;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddScoped<IFondBurseService, FondBurseService>();
builder.Services.AddScoped<IFondBurseMeritRepartizatService, FondBurseMeritRepartizatService>();

builder.Services.AddControllers();
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
builder.Services.AddDbContext<BurseDBContext>(options => options.UseSqlServer(builder.Configuration.GetConnectionString("BurseConnectionStrings")));
var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();

app.UseAuthorization();

app.MapControllers();

app.Run();
