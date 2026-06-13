# Agent-Excel

A local HTTP API service accompanied by a lightweight CLI client designed to control Microsoft Excel via RESTful APIs, specifically tailored for AI Agents.

## Core Features

- **High-Frequency Control**: Enables AI agents to read and write Excel sheets currently open and operated by the user at extreme speeds.
- **Client-Daemon Architecture**:
  - **Daemon**: A persistent background service holding the Excel COM object handles long-term, exposing local HTTP endpoints.
  - **Client**: A lightweight CLI trigger to start, stop, or check the status of the daemon.
- **COM Safety**: Focuses on attaching to existing Excel instances and preventing memory leaks or zombie processes.

## Tech Stack

- **Language/Framework**: C# 13 / .NET 10 (Windows platform only)
- **Framework**: ASP.NET Core Minimal APIs
- **Dependency**: Microsoft Office Excel COM Interop

## Quick Start

### CLI Lifecycle Management

1. **Start the Service** (Runs the daemon silently in the background):
   ```bash
   AgentExcel.exe start
   ```
2. **Check Status**:
   ```bash
   AgentExcel.exe status
   ```
3. **Stop the Service**:
   ```bash
   AgentExcel.exe stop
   ```

### API endpoints (Daemon Side)

Once the daemon starts, it exposes HTTP endpoints returning YAML/Plain Text:
- `GET /status` - Returns daemon and Excel attachment health status.
- `POST /exit` - Gracefully releases COM resources and shuts down the service.

## Developer Guide

For detailed guidelines, code style standards, and COM memory management rules, please refer to [AGENTS.md](AGENTS.md).
