# Use the .NET 6 SDK image for build and publish
FROM mcr.microsoft.com/dotnet/sdk:6.0 AS build

# Set the working directory inside the container
WORKDIR /src

# Copy only the project file and restore dependencies
COPY NPOIGenerateExcelScatterChart.csproj ./
RUN dotnet restore

# Copy the rest of the application source code to the working directory
COPY . ./

# Build the application in Release mode
RUN dotnet build -c Release

# Publish the application, this step creates the self-contained app
RUN dotnet publish -c Release -o /app/publish

# Use the runtime image for a smaller final image
FROM mcr.microsoft.com/dotnet/runtime:6.0

# Set the working directory for the app in the final image
WORKDIR /app

# Install fonts for NPOI (to handle font issues)
RUN apt-get update && apt-get install -y fontconfig \
    && apt-get clean

## Copy the compiled output from the build image
#COPY --from=build-env /src/bin/Release/net6.0 /app

# Copy the published output from the build stage
COPY --from=build /app/publish .

# Set the entry point for the application
ENTRYPOINT ["dotnet", "NPOIGenerateExcelScatterChart.dll"]
