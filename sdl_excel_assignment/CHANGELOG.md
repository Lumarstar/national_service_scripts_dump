# Changelog

All notable changes to this project will be documented in this file.

## Version Log

### [v1.2.0] - 6 Dec 2023
- Removed image insertion functions due to AppScript's incompatibility with Google Drive. Instead, the assignment completion message was edited 
so that trainees will have to submit a screenshot with it included as proof (as the message is not transferrable outside each script run).
- Was not able to replicate the bar chart bug. Tried various combinations of the bar chart title.

### [v1.1.0] - Date unknown
- Functionality to prevent cheating was implemented, namely the hidden tracker sheet and an image that will be displayed upon completion of the assignment.
- Piloted this version during batch 41/23 on 22 Nov 2023.
  - Discovered a bug where the image (from Google Drive) is not retrieved properly from the provided URL.
  - Discovered another bug where the bar chart title needs a very specific font and formatting.

### [v1.0.0] - Date unknown
- Functionality to check assignment implemented using Google Sheets.

### [v0.0.0] - Date unknown
- Functionality to check assignment was implemented on [an Excel workbook](https://github.com/tewenhao/national_service_scripts_dump/blob/8c910f7824ead88923832f90acbc42815e7e588c/sdl_excel_assignment/ASA_SDL_Automated_Checks.xlsm)
using VBA Macros.

[v1.2.0]: #v120---6-dec-2023
[v1.1.0]: #v110---date-unknown
[v1.0.0]: #v100---date-unknown
[v0.0.0]: #v000---date-unknown
