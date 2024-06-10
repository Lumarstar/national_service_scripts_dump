# SDL Excel Assignment

Scripts to automate the marking of a Microsoft Excel assignment that serves as the culmination of the Administrative Support Assistant trainees' Microsoft Excel
module, implemented on Google Sheets using Google App Script.

## Setup Guide

### On Master Sheet

1. Access the master copy of the Google Sheet through
[this link](https://docs.google.com/spreadsheets/d/1IPgtKAcr2ESv4UnOSGEeMlch_MSFvMLA9TrPxDCw2Ns/edit#gid=524336585).

2. You should be able to see this page loaded up:
<div align='center'>
  <img src="https://github.com/tewenhao/national_service_scripts_dump/assets/63058663/e1e8f221-c9b4-480c-b576-66d10e3c0c42" alt="image showing landing page of master sheet" width="70%" height="70%">
</div>

3. Make a copy of the master sheet. Allow the associated app script to be copied over as well. **All subsequent steps will be on the copied sheet.**
<div align='center'>
  <img src="https://github.com/tewenhao/national_service_scripts_dump/assets/63058663/cbb3bb77-5854-4f8e-8618-a4922f00f039" alt="image showing where to find 'Make a copy'" width="70%" height="70%">
</div>
<div align='center'>
  <img src="https://github.com/tewenhao/national_service_scripts_dump/assets/63058663/dd348ccb-0d1e-4e56-8f62-39605d85c8b2" alt="image showing coping options" width="70%" height="70%">
</div>

### On Copied Sheet

4. Click the `GRADE WORK` button and run the programme. Since this (should) be the first time, you would need to grant the programme permission to access
your sheet. A separate window should pop up.

    - Press `OK` here
      <div align='center'>
        <img src="https://github.com/tewenhao/national_service_scripts_dump/assets/63058663/3e1ee77a-810f-4e5a-b67f-83b9356a6b9b" alt="image showing 'authorisation required to run programme'" width="70%" height="70%">
      </div>
    - Proceed to sign in with your Google Account
      <div align='center'>
        <img src="https://github.com/tewenhao/national_service_scripts_dump/assets/63058663/45f4be0c-f094-48b8-b19d-afc7e5101464" alt="image showing 'Google Account Sign in Screen'" width="70%" height="70%">
      </div>
    - Click `Advanced`, then `Go to SDL Automated Assignment (unsafe)`
      <div align='center'>
        <img src="https://github.com/tewenhao/national_service_scripts_dump/assets/63058663/a34760b3-fbdf-4bd1-bfe0-0930bc7e7be5" alt="image showing where is 'Advanced'" width="70%" height="70%">
      </div>
      <div align='center'>
        <img src="https://github.com/tewenhao/national_service_scripts_dump/assets/63058663/889add6c-fdf5-4b2a-b7ab-1c278f588bb0" alt="image showing where is 'Go to SDL Automated Assignment (unsafe)'" width="70%" height="70%">
      </div>
    - Allow the programme to access your Google Account
      <div align='center'>
        <img src="https://github.com/tewenhao/national_service_scripts_dump/assets/63058663/4461ee0a-c988-4f0f-a7e8-1d77788344f2" alt="image showing where is 'Go to SDL Automated Assignment (unsafe)'" width="70%" height="70%">
      </div>

5. At this point, the window should have been closed. Now, you should be able to use the sheet as intended.

## Usage Guide

Following [this guide here](https://github.com/tewenhao/national_service_scripts_dump/blob/3c5b1c8510e3934e52ab69d78c90c778b8bbabbd/sdl_excel_assignment/ASA%20SDL%20Excel%20Assignment%20Instructions.pdf),
you should be able to complete the assignment.

When grading, there are a few possible outcomes for each task:
- If the task is completed correctly, a pop-up window will be shown indicating the task is completed correctly, and `Completed` will be filled in in the
corresponding cell in the table.
- If the task is not completed correctly, a pop-up window will show which cell requires correction. (But do know that this is usually caused by an incorrect
formula anyway.)
- If the task is not completed but somehow, `Completed` has been filled in, the programme will remove `Completed` from the table starting from the cell that
corresponds to said task and all tasks after it.

## Change Log

Refer to [CHANGELOG.md]() for the change log!

## Acknowledgement

I would like to thank the following people:

- 41/23 C2306 Salmaan is recognized for his contributions to the script, particularly for his efforts in resolving a bug that hindered the correct display of an image, and for identifying another bug necessitating specific fonts and formatting for the title of the bar chart.

## Copyright

Copyright © 2023 - 2024 Tew En Hao

This repository is authored by Tew En Hao, ASCC 57/23, BMTC SCH V Coyote Company Platoon 2 (38/23 to 41/23) and Platoon 3 (42/23 onwards) Section
Commander/Assistant Trainer.

All Rights Reserved.
