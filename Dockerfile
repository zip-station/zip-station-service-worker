FROM mcr.microsoft.com/dotnet/sdk:10.0 AS build
WORKDIR /src

COPY ZipStation.Worker/ZipStation.Worker.csproj ZipStation.Worker/
RUN dotnet restore ZipStation.Worker/ZipStation.Worker.csproj

COPY . .
RUN dotnet publish ZipStation.Worker/ZipStation.Worker.csproj -c Release -o /app/publish --no-restore

FROM mcr.microsoft.com/dotnet/aspnet:10.0 AS runtime
WORKDIR /app
COPY --from=build /app/publish .

ENTRYPOINT ["dotnet", "ZipStation.Worker.dll"]
