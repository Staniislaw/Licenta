# Etapa 1: Build
FROM mcr.microsoft.com/dotnet/sdk:6.0 AS build
WORKDIR /app

# Copiaza fi?ierul .csproj ?i restore dependen?ele
COPY *.csproj ./
RUN dotnet restore

# Copiaza restul codului
COPY . ./
RUN dotnet publish -c Release -o out

# Etapa 2: Runtime
FROM mcr.microsoft.com/dotnet/aspnet:6.0
WORKDIR /app
COPY --from=build /app/out ./

# Asculta pe portul pe care îl ofera Render
ENV ASPNETCORE_URLS=http://+:$PORT
EXPOSE 10000

# Înlocuie?te cu numele real al fi?ierului tau .dll
ENTRYPOINT ["dotnet", "MyApi.dll"]
