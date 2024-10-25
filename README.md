# Project Title: Gate Management Automation
Automation of an Excel-based document tracking critical flight information, including flight numbers, departure/arrival airports, scheduled times (STD/STA), passenger numbers, and special requirements


## Background and Overview
This project aims to automate the management of gate assignments and related tasks for an airline's operational workflow. Previously, managing gate assignments, prioritizing flights, and handling swaps was a time-consuming and manual process that often led to errors and inefficiencies. This automation streamlines these processes by utilizing VBA macros in Excel to handle repetitive tasks efficiently, allowing for better time management and operational accuracy.

The main goals of this project include:
- Reducing the time spent on gate sheet building and management tasks.
- Minimizing errors associated with manual data entry and updates.
- Enhancing the overall workflow of the operational staff by providing a user-friendly interface for data manipulation.

## Data Structure Overview
The data involved in this project consists of several key worksheets that hold relevant information for gate management operations. Below is a brief overview of the data structure:

### Worksheets


1. **Today's Gate Sheet (1)**:
   - **Columns**:
     - Column A: Departure airport code for each flight.
     - Column B: Flight numbers for today's operations.
     - Column C: Unique aircraft number to identify which aircraft will be used for each flight.
     - Column D: Gate assignment for each flight.
     - Column E: Arrival airport code for each flight.
     - Column F: # of passengers on each flight out of 186.
     - Column G: Time of departure for each flight. Formatted as military time (hhmm).
     - Column H: Time of arrival for each flight. Formatted as military time (hhmm).
     - Column I: Column used to correctly label Charter flights as either "CARGO" flights or "FERRY" flights.
     - Column J: Column used to input the actual departure time of flights leaving MSP and also used to input ETA times when inbound flights to MSP go airborne.
     - Column K: Column used to input the actual arrival (block-in) time of flights arriving into MSP.
     - Column L: # of passengers on each flight.
     - Column M: Column used to correctly label whether each flight is an ORIG, TURN, or TERM.
     - Column N: Area for inputting notes or extra important information for each flight. e.g. "INOP APU", "PRIORITY", "LAV/WATER".
       
![image](https://github.com/user-attachments/assets/4996e055-be6a-42f3-9729-8c1137caf622)

2. **S&P**:
   - **Columns**:
     - Column C: Flight numbers (including multiple formats)
     - Column D: Corresponding source values related to the flights
     - Column I: Additional identifiers (typically 3 to 4-digit numbers) for priority management

![image](https://github.com/user-attachments/assets/40014ab6-ecdb-41c8-b7f1-32bcc623473a)

This structure allows for quick lookups and updates based on the flight identifiers, ensuring that all relevant data is interconnected for effective management.

## Executive Summary
The Gate Management Automation project leverages Excel VBA macros to enhance the efficiency of gate management operations within an airline setting. By automating the processes of gate assignment updates, priority management, and swap handling, this project reduces manual labor and the potential for human error.

### Key Features
- **Gate Change Management**: Automates the process of changing gate assignments by looking up flight identifiers and updating the corresponding gate information seamlessly.
- **Priority Assignment**: Allows users to add "PRIORITY" tags to specific flights easily, ensuring that critical operations are highlighted and addressed promptly.
- **User-Friendly Interface**: The macro prompts users for confirmation before executing actions, making it easier for operational staff to interact with the tool and implement changes with confidence.

### Impact
By implementing this automation, operational staff can focus on more strategic tasks rather than getting bogged down in repetitive, time-consuming processes. This project ultimately leads to improved operational efficiency, better resource allocation, and enhanced service delivery within the airline.

