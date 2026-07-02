
# docs2pdfs

A lightweight Windows context menu tool that converts `.docx` files to PDF with a single right-click.

## How It Works

After installation, right-click any `.docx` file in Windows Explorer and select **"Convert to PDF"**. The PDF is created in the same directory as the source file.

## Requirements

- Windows 10/11
- Microsoft Word (uses Word COM Automation for conversion)
- .NET Framework 4.7.2+

## Installation

1. Download `DocxToPdf_Setup.exe` from [Releases](../../releases)
2. Run the installer (requires admin privileges)
3. Done — the context menu entry is added automatically

Uninstalling via Add/Remove Programs removes all files and registry entries.

## Build from Source

1. Open `docs2pdfs.slnx` in Visual Studio
2. Build in **Release** mode
3. Compile `setup/docs2pdfs.iss` with [Inno Setup](https://jrsoftware.org/isinfo.php) to generate the installer

## Project Structure

```
docs2pdfs/
├── Program.cs        # Entry point, argument handling
├── converter.cs      # Word Interop → PDF conversion
└── setup/
    └── docs2pdfs.iss # Inno Setup installer script
```

## License

MIT
