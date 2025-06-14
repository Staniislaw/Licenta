﻿using Burse.Data;
using Burse.Services.Abstractions;
using Burse.Services;

using Microsoft.EntityFrameworkCore;
using Burse.Helpers;
using QuestPDF.Infrastructure;

var builder = WebApplication.CreateBuilder(args);
// Add services to the container.
builder.Services.AddScoped<IFondBurseService, FondBurseService>();
builder.Services.AddScoped<IFondBurseMeritRepartizatService, FondBurseMeritRepartizatService>();

// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
builder.Services.AddDbContext<BurseDBContext>(options => options.UseSqlServer(builder.Configuration.GetConnectionString("BurseConnectionStrings")));
builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowAll", builder =>
        builder.AllowAnyOrigin()
               .AllowAnyMethod()
               .AllowAnyHeader());
});
builder.Services.AddControllers().AddJsonOptions(options =>
{
    options.JsonSerializerOptions.ReferenceHandler = null; // sau nu seta deloc această opțiune

});

builder.Services.AddScoped<IStudentService, StudentService>();
builder.Services.AddScoped<IBurseIstoricService, BurseIstoricService>();
builder.Services.AddScoped<GrupuriDomeniiHelper, GrupuriDomeniiHelper>();
builder.Services.AddScoped<IPdfGeneratorService, PdfGeneratorService>();
builder.Services.AddScoped<IGrupuriService, GrupuriService>();
builder.Services.AddScoped<StudentExcelReader, StudentExcelReader>();
builder.Services.AddSingleton<AppLogger>();

QuestPDF.Settings.License = LicenseType.Community;

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();

app.UseCors("AllowAll");

app.UseAuthorization();

app.MapControllers();

app.Run();
