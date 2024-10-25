# Gate Sheet Management Automation
Automation of Excel-based documents tracking critical flight information, including flight numbers, departure/arrival airports, scheduled times (STD/STA), passenger numbers, and special requirements.

## Background and Overview
This project aims to automate the management of gate assignments and related tasks for an airline's operational workflow. Previously, managing gate assignments, prioritizing flights, and handling swaps was a time-consuming and manual process that often led to errors and inefficiencies. This automation streamlines these processes by utilizing VBA macros in Excel to handle repetitive tasks efficiently, allowing for better time management and operational accuracy.

**The main goals of this project include:**
- Reducing the time spent on gate sheet building and management tasks.
- Minimizing errors associated with manual data entry and updates.
- Enhancing the overall workflow of the operational staff by providing a user-friendly interface for data manipulation.

## Data Structure Overview
The data involved in this project consists of several key worksheets that hold relevant information for gate management operations. Below is a brief overview of the data structure:

### Worksheets


1. **Gate Sheet (1)**:
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
       
![image](https://github.com/user-attachments/assets/99568d94-dd7f-4708-8450-e4c08c1fef77)


2. **Swap Sheet (S&P)**:
   - **Columns**:
     - Column A: Unique aircraft number to identify which aircraft will be used for that flight.
     - Column B: The "WILL OPERATE" column will visually help with which aircraft will operate which flight. No rules are run in this column.
     - Column C: Flight numbers (including multiple formats)
     - Column D: New gate value if that flight operates out of a new gate.
     - Column I: Flight #'s that will have "PRIORITY" for that specific flight. 

![image](https://github.com/user-attachments/assets/40014ab6-ecdb-41c8-b7f1-32bcc623473a)


**Gate Plot**:
   - **Columns**: Contain flight gate data, shaded according to flight type, allowing quick identification of gate assignments and flight types.
   - **Rows**: Contain time intervals starting at 5 a.m. and ending at 1 a.m., with 5-minute intervals between each hour.
   - **Purpose**: A visual representation of the gate sheet with color-coded data to indicate the types of flights departing from each gate.
   - **Color Coding**:
     - ORIG (origin flights) are shaded green. TURNS (turnaround flights) are shaded purple, and TERMS (terminating flights) are shaded light blue. These flights are visually distinguished to streamline gate management.
       
![image](https://github.com/user-attachments/assets/bd53f073-5a52-4c19-a66a-fc5450448629)
  

This structure allows for quick lookups and updates based on the flight identifiers, ensuring that all relevant data is interconnected for effective management.

## Executive Summary
The Gate Management Automation project leverages Excel VBA macros to enhance the efficiency of gate management operations within an airline setting. By automating the processes of gate assignment updates, priority management, and swap handling, this project reduces manual labor and the potential for human error.

**Gate Sheet Macro**:
- **Sub gatesheet1()**: This macro cleans up and formats the gate sheet data:
     - Font Formatting: Sets the entire sheet’s font to Calibri, size 20, and bold.
     - Clearing Contents: Clears the contents in columns H and K from row 3 downwards.
     - Data Rearrangement: Cuts data from column J and pastes it into column H.
     - Border Styling: Adds thin black borders to all edges and inside lines of the data range from columns A to N, starting from row 3 down to the last used row.
     - Time Format: Sets columns G and H to display in "HHMM" time format.
     - Deleting Extra Columns: Deletes columns O to Q up to row 300.
     - Fill Color in Column N: Fills cells in column N with white background color from row 3 to the last row with data.
       
- **Sub gatesheet2()**: Continues with formatting the gate sheet along with inputting specific flight information:
  - PAX Count Formatting:
     - Cleans up values in column F by removing "0/" prefixes, formats the column as text, copies its values to column L, and appends "/186" to each non-empty cell in column F.
  - Data Formatting and Styling:
     - Adds borders around the data range in columns A to N.
     - Sets the font style in columns G and H to Calibri, size 20, bold.
     - Formats columns G and H to "HHMM" time format, then deletes columns O to Q.
  - Adding Customs Information:
     - For cells in column C containing "N", it populates columns D, I, and N with predefined values ("WC", "CARGO", "AMZ / CREW RIDE").
     - If column B contains numeric values greater than or equal to 8000, a specific value is inserted in column I.
  - Customs Classification:
     - Checks if values in column A match either customs values (like "AUA", "BZE") or pre-cleared values (like "YYZ", "YVR"). Depending on the match, it labels column N as "CUSTOMS" or "CUSTOMS PRE-CLEARED."
  - Font Adjustments by Cell Color: For cells in columns G and H with specific fill colors, the macro modifies font boldness and size.
  - Remove Letters in Column C: Strips "N" and "A" from values in column C.
    
- **Sub gatesheet3()**: This macro handles the cutting and moving of each outbound flight whose tail value matches an inbound flight. This signifies that the outbound flight will turn off the inbound aircraft.
  - Set Up:
     - Identifies the worksheet and the last row in column C.
     - Initializes a collection to hold rows that need moving.
  - Match and Move Rows:
     - Iterates backward through column C, looking for rows with a white fill color.
     - When a white-filled row is found, the macro searches above it for a row with the same value in column C but with a different fill color.
     - If a match is found, it moves the white-filled row directly below the matched row, ensuring the rows with the same value are grouped together.
  - Cleanup:
     - Deletes rows with empty values in column A to tidy up the data.
  - Finalization:
     - Shows a message box indicating that the "Tails Matching" process is complete and advises running the macro three times for accurate results.

- **Sub gatesheet4()**: This macro categorizes flight data in Column C based on specific rules, specifically identifying "TURN," "TERM," and "ORIG" flights.
  - Loop Through Column C:
     - It checks whether each cell in Column C (excluding the last 20 rows) has a background fill color other than white.
     - If the current cell value matches the cell value in the row below, it inserts "TURN" in Column M of that row.
     - If the current cell value differs from the cell above it, it inserts "TERM" in Column M.
  - Second Loop (For White Fill Cells):
     - It checks cells with a white background fill.
     - If the cell’s value differs from the above row, "ORIG" is inserted into Column M.
  - Purpose:
     - TURN: Likely indicates a turnaround flight (same flight number or aircraft returns quickly).
     - TERM: Could signify the termination of a flight or a change in aircraft.
     - ORIG: Marks the origin point or starting flight in a sequence.
   
- **Sub gatesheet5()**: Organizes flight rows based on morning flight times (found in Column G) and adjusts the order of flights that have an ORIG status (with a white fill in Column C).
  - Loop to Identify Rows with White Fill in Column C:
     - Starting from the last row, the code looks for cells in Column C with a white fill (indicating an "ORIG" flight).
     - For each row found, the flight time in Column G is compared.
     - It searches upwards through Column G to find the next available white-fill cell with a smaller flight time.
  - Move Row:
     - Once it finds a suitable row with a smaller flight time, it moves the current row below it.
     - The Cut operation is used to shift the row into the correct position.
  - Clean Up:
     - After organizing the rows, it checks Column A and deletes any empty rows.
   
- A powerpoint breakdown of the Gate Sheet macros' key steps and logic can be downloaded [here](https://github.com/Shepshub/Aviation-Gate-Sheet-Automation/raw/refs/heads/main/Gate%20Sheet%20Macro%20Final.pptx) and the actual VBA code can be downloaded [here](https://github.com/Shepshub/Aviation-Gate-Sheet-Automation/raw/refs/heads/main/Sheet%20Gate%20Macro.docx).
- A demo video on how to use the Gate Sheet macro can be viewed [here](https://drive.google.com/file/d/1J36SXwEwJ_k68KOaRN1xk1t_bXjyBvG4/view?usp=drive_link).

**Swap Sheet Macro (S&P)**:
- **Sub xTAILSWAP()**: Identifies and transfers aircraft information between the Swap Sheet and the Gate Sheet based on matching flight numbers while also formatting the output with borders for clarity.
  - Setup and Prompt:
     - Sets references to the two worksheets and prompts the user to confirm if they want to proceed with the tail swap operation.
  - Loop and Split Values:
     - For each cell in column C of the "S&P" sheet, splits the cell’s value by common delimiters (commas, slashes, or hyphens) to isolate numbers.
     - Removes any non-numeric characters to get a clean number string.
     - Search and Paste Matching Values:
  - Searches for each cleaned value in column B of the "1" sheet.
     - If a match is found, copies the corresponding value from column A of the same row on the "S&P" sheet and pastes it in column C of the matching row on the "1" sheet, adding a thick border around the pasted cell.
  - Copy Additional Data:
     - If column D in the "S&P" sheet contains data for the current row, this value is copied and pasted into column D of the same matching row on the "1" sheet, also with a thick border.
  - Completion Message:
     - Displays a "Done" message if the operation completes successfully or “Okay” if the user chooses not to proceed.

- **Sub xGATECHANGE()**: Identifies and transfers gate information between the Swap Sheet and the Gate Sheet based on matching flight numbers while also formatting the output with borders for clarity.
  - Setup and Prompt:
     - Sets up references to the worksheets and prompts the user to confirm if they want to proceed with changing the gate information.
  - Loop and Split Values:
     - Iterates through each cell in column C of the "S&P" (Swap) sheet, splitting cell values based on commas, backslashes, or hyphens to get individual numbers.
     - Removes non-numeric characters from each value to obtain a clean number string.
  - Search and Paste Matching Gate Information:
     - Searches for the cleaned value in column B of the "1" (Gate) sheet.
     - If a match is found, retrieves the corresponding gate information from column D of the same row in the "S&P" sheet and pastes it into column D of the matching row in the "1" sheet.
     - Thick borders around the pasted cell are added to make it visually distinct.
  - Completion Message:
     - Displays a "Done" message if successful or an “Okay” message if the user chooses not to proceed.

- **Sub xPRIORITY()**: Identify and add the term "PRIORITY" to specific flights in the Gate Sheet based on matching flight numbers from the Swap Sheet.
  - Setup and User Prompt:
     - Establishes references to the "S&P" (Swap) sheet and the "1" (Gate) sheet.
     - Prompts the user with a message box asking if they want to add "PRIORITY" to matched rows.
  - Exit Condition:
     - If the user selects "No," the macro exits without making any changes.
  - Loop Through Values:
     - Iterates through each cell in column I of the "S&P" sheet, starting from row 3.
     - Extracts 3 to 4-digit numeric values from each cell, creating a cleaned version of the value.
  - Search for Matches:
     - Searches for the cleaned value in column B of the "1" sheet.
     - If a match is found, it checks the corresponding cell in column N of the "1" sheet.
  - Adding "PRIORITY":
     - If column N already has a value, it appends " / PRIORITY" to the existing value.
     - If the cell is empty, it simply sets the value to "PRIORITY."

- A powerpoint breakdown of the Swap Sheet macros' key steps and logic can be downloaded [here](https://github.com/Shepshub/Aviation-Gate-Sheet-Automation/raw/refs/heads/main/Swap%20Sheet%20Macro%20Final.pptx) and the actual VBA code can be downloaded [here](https://github.com/Shepshub/Aviation-Gate-Sheet-Automation/raw/refs/heads/main/Swap%20and%20Priority%20Sheet%20Macro.docx).
- A demo video on how to use the Swap Sheet macro can be viewed [here](https://drive.google.com/file/d/1rBiw-O2IZo8goLTl7HjTiOuk5Uc_6wRy/view?usp=drive_link).

**Gate Plot Macro**:

- Sub Plot1() visualizes and plots flight gate information based on time intervals and gate numbers. The key idea is to match flight times to specific time slots on the gate plot, ensure proper spacing based on gate usage, and include additional flight information (like destination and aircraft type) in the appropriate places.
- Sub Plot2() macro fills in cells in the gate plot worksheet based on inbound flight data from the gate sheet. It finds flights marked with "TURN," rounds the arrival time if needed to the nearest 5 minutes, finds the appropriate gate, and shades cells purple. It also fills and transfers specific gate information like airport codes and flight numbers into the purple-marked cells.
- Sub Plot3() subroutine handles the shading of the remaining outbound flights in the gate plot. These remaining flights originate from MSP. The macro locates the remaining green-filled cells with a 3 to 4-letter value and shades nine cells to the left with the same green color, signifying the amount of ground time given to fully load/board and push the flight out.
- Sub Plot4() works similarly to the Sub Plot2() macro, although, instead of looking for TURN value flights, it searches for TERM value flights, signifying that the flight will terminate to the hangar and not turn back out. It extracts the airport code, flight number, and gate assignment from the gate sheet and populates the correct cells in the gate plot. Then, shading nine cells to the right of the arrival time signifying the time needed to fully deplane, unload bags, and get the aircraft off of the gate.

- A powerpoint breakdown of the Gate Plot macros' key steps and logic can be downloaded [here](https://github.com/Shepshub/Aviation-Gate-Sheet-Automation/raw/refs/heads/main/Gate%20Plot%20Macro%20Final.pptx) and the actual VBA code can be downloaded [here](https://github.com/Shepshub/Aviation-Gate-Sheet-Automation/blob/main/Gate%20Plot%20Macro.docx).
- A demo video on how to use the Gate Plot macro can be viewed [here](https://drive.google.com/file/d/1y6AU7GrpoWOAUvuk7VYGMusH5IDccW2N/view?usp=drive_link). Also a video on how to display inbound ETA's effectively can be viewed [here](https://drive.google.com/file/d/1pwc2dn-Yz74o74bhURSSjouanFo09Ewe/view?usp=drive_link).

### Key Features
- **Gate Sheet Building and Management**: Automates the process of building gate sheets and changing aircraft routes and gate assignments assignments by looking up flight identifiers and updating the corresponding aircraft and gate information seamlessly.
- **Priority Assignment**: Allows users to add "PRIORITY" tags to specific flights easily, ensuring that time critical operations are highlighted and addressed promptly.
- **User-Friendly Interface**: The macro prompts users for confirmation before executing actions, making it easier for operational staff to interact with the tool and confidently implement changes.

### Impact
By implementing this automation, operational staff can focus on more strategic tasks rather than getting bogged down in repetitive, time-consuming processes. This project ultimately leads to improved operational efficiency, better resource allocation, and enhanced service delivery within the airline.

