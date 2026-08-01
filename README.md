# SKAR_specs

This app is a response to a very bad situation with computer specification software around the internet. Lots of adware, lots of paid software. Lots of hard to use software. Not even Windows itself have a utility that shows all info in one place. Well, task manager probably does the best job at that :) (newer windows builds have half baked about screen, better than used to be). So this program generates a HTML report of computer specifications. Double click a file, wait a little, open the file. Read all necessary information about computer in one page. Take the file forward, send, take a photo of summary page. 

SKAR_specs is a Windows system-inventory utility that collects hardware, operating-system, network, printer, software, and license information into a portable report. It produces HTML by default and can also export JSON or plain text.

> [!WARNING]
> Reports may contain sensitive information, including usernames, device serial numbers, MAC/IP addresses, network shares, installed software, license status, and remote-support IDs. Review every report before sharing it or attaching it to a public issue.

## Requirements

- Windows 10 or Windows 11
- .NET Framework 4.6 or a later compatible .NET Framework 4.x runtime
- Visual Studio 2022 with .NET desktop build tools, or equivalent MSBuild tooling, to build from source

The application is Windows-only because it uses WMI, the Windows Registry, and Windows command-line utilities.

## Build

Open `SKAR_specs.sln` in Visual Studio and build the `Release` configuration, or run:

```powershell
.\Build.ps1
```

The build script defaults to `Release`, verifies the expected output files, and prints the executable's SHA-256 checksum. To invoke MSBuild directly instead, run:

```powershell
dotnet msbuild .\SKAR_specs.sln -p:Configuration=Release
```

The executable is written to `bin\Release\SKAR_specs.exe`. The project has no third-party package dependencies.

## Usage

Run `SKAR_specs.exe`. A timestamped report is created beside the executable.

```powershell
# HTML (default)
.\SKAR_specs.exe

# JSON
.\SKAR_specs.exe --json

# Plain text
.\SKAR_specs.exe --text

# Include timing and diagnostic details
.\SKAR_specs.exe --debug
```

Options are case-insensitive. The short Windows-style forms `-json`, `/json`, `-txt`, `/txt`, `-text`, `/text`, `-debug`, and `/debug` are also accepted.

## What it collects

- System summary, Windows version, TPM, and current-user details
- Motherboard, BIOS, processor, memory, disks, graphics, and monitors
- Network adapters, IP configuration, shares, mapped drives, and connectivity
- Printers and printer ports
- Installed desktop software and UWP apps
- Windows/Office license status and selected remote-support IDs

For connectivity reporting, SKAR_specs sends HTTPS requests to `ifconfig.me`, `api.ipify.org`, and `ipinfo.io` to determine the public IP address. It also pings `8.8.8.8`, `1.1.1.1`, and `9.9.9.9`. No report file is uploaded by SKAR_specs.

## Contributing

Contributions are welcome; see [CONTRIBUTING.md](CONTRIBUTING.md). Never include an unredacted SKAR_specs report or other sensitive inventory data in a public issue.

## License

SKAR_specs is free software licensed under the [GNU General Public License v3.0 only](LICENSE) (`GPL-3.0-only`).
