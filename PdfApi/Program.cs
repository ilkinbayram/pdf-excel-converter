using System.Text;
using PdfApi.Service;

var builder = WebApplication.CreateBuilder(args);

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

// Add services to the container.
builder.Services.AddControllers();
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

// builder.Services.AddScoped<IFileHandlerService, SpireFileHandlerService>();
// builder.Services.AddScoped<IFileHandlerService, VtbFileHandler>();
builder.Services.AddScoped<IFileHandlerService, ByteScoutFileHandler>();

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
