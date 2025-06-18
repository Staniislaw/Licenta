FROM mcr.microsoft.com/dotnet/sdk:6.0 AS build
WORKDIR /app

# Copiaza doar fi?ierul .csproj din folderul Burse
COPY Burse/*.csproj ./Burse/
WORKDIR /app/Burse
RUN dotnet restore

# Copiaza restul codului
COPY . .

RUN dotnet publish -c Release -o out

FROM mcr.microsoft.com/dotnet/aspnet:6.0
WORKDIR /app
COPY --from=build /app/Burse/out ./

ENV ASPNETCORE_URLS=http://+:$PORT
EXPOSE 10000

ENTRYPOINT ["dotnet", "NumeleProiectului.dll"]
