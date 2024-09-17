# NPOI Excel Scatter Chart Report

This project demonstrates how to generate an Excel report with scatter charts using the NPOI library in a .NET 6 console application. The project builds and runs within a Docker container.

## Table of Contents

1. [Project Overview](#project-overview)
2. [Prerequisites](#prerequisites)
3. [Getting Started](#getting-started)
4. [Code Explanation](#code-explanation)
5. [Docker Setup](#docker-setup)
6. [Build and Run](#build-and-run)
7. [Troubleshooting](#troubleshooting)

---

## Project Overview

This project uses the NPOI library to generate Excel files with formatted data and scatter charts. NPOI provides a .NET interface to work with Excel files similar to Apache POI (for Java). The project outputs an Excel file with a scatter chart based on some dummy data (related to football in this example).

## Prerequisites

- **Docker** installed on your machine.
- **.NET 6 SDK** (for local development).
- Basic understanding of .NET applications and Docker.

---

## Getting Started
Run the project locally (optional): You can run the project locally before containerizing it. You need the .NET 6 SDK installed.


dotnet build
dotnet run
Run using Docker: You can use the Dockerfile provided to build and run the project in a Docker container.

Code Explanation
The core of the project is written in a single file Program.cs. Here's a breakdown:

NPOI Library: This library is used to generate Excel reports without requiring Excel to be installed. The scatter chart and data are written directly to an Excel file.

Main Components:

Data Generation: The code simulates football data with random values for the scatter chart.
Chart Creation: A scatter chart is created using the generated data.
Excel Formatting: The Excel sheet is formatted with headers and styled cells for a clean presentation.
Key Sections:
CreateData Method: Generates random data for the scatter chart.
CreateScatterChart Method: Plots the scatter chart using NPOI based on the generated data.
FormatHeader Method: Applies formatting and styling to the header rows.
Docker Setup
The project is containerized using Docker for easy deployment and execution in any environment. The Dockerfile is divided into two stages:

Build Stage: Uses the .NET SDK to build the project.
Runtime Stage: Uses the .NET Runtime to execute the built application.
Dockerfile Breakdown:
Stage 1 - Build the Application:

A .NET 6 SDK image is used to build the project.
The project files are copied, and NuGet dependencies are restored using dotnet restore.
The project is built in Release mode.
Stage 2 - Runtime Environment:

A .NET Runtime image is used to run the compiled .NET application.
Fonts are installed to handle the NPOI font requirements.
The build artifacts are copied from the build stage, and the application is run.
Build and Run
Build the Docker Image
To build the Docker image, run the following command:


docker build -t football-report .
Run the Docker Container
Once the image is built, you can run it using the following command:


docker run --rm -it football-report
The output Excel file will be generated inside the container. You can mount a volume to access the generated file from your host system.

To run with volume mounting and export the Excel file to your local system:


docker run --rm -it -v $(pwd)/output:/app/output football-report
This will create an output directory in your current directory and place the generated Excel report there.

Troubleshooting
Missing Fonts Error
If you encounter the error:


Unhandled exception. SixLabors.Fonts.FontException: No fonts found installed on the machine.
It means the container lacks the necessary fonts for NPOI. The Dockerfile includes a line to install fontconfig for handling this.

Make sure to use the updated Dockerfile that includes this:
RUN apt-get update && apt-get install -y fontconfig && apt-get clean
Fallback Package Error
If you encounter an error related to missing NuGet packages, ensure that the following command in the Dockerfile restores packages correctly:

Conclusion
This project demonstrates how to generate Excel reports with charts in a .NET 6 console application, containerized using Docker. The example can be extended to different use cases, such as generating reports based on real data.
