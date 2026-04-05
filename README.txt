ARGUS // Web-Only Source Package

What changed:
- Web shell is now the only intended UI path.
- WinForms UI files are excluded from the project build.
- Assembly output name is now Argus.
- Convert tab is wired to the backend.
- Media tab settings now affect rename behavior.
- Image metadata rename support was added for supported image types.
- Conversion support:
  - Built-in image conversion: JPG / PNG
  - ffmpeg when installed: MP3 / WAV / MP4
  - LibreOffice when installed: document -> PDF
  - ImageMagick when installed: image -> PDF
- Undo now deletes generated conversion outputs.

Included for build stability:
- Local direct references to UglyToad PdfPig DLLs are used in the csproj.

Expected build path on Windows:
1. Open terminal in this folder.
2. dotnet publish FileOrganizer.csproj -c Release -r win-x64 --self-contained false -o .\publish
3. Run publish\Argus.exe

Notes:
- No rebuilt EXE was produced here because this environment does not have dotnet.
- Convert tab will honestly report capability depending on what is installed on the target machine.
