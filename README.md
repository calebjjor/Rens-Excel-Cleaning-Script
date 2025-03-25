# Customer Orders Filtering Tool
A Python script to clean Excel spreadsheets based on the desired period of time.

## Downloading and Setting Up `script.exe`

### Step 1: Download `script.exe` from GitHub
1. **Go to the GitHub Repository**
   - Open a web browser and navigate to the GitHub repository where `script.exe` is hosted.

2. **Download the File**
   - Click on the `script.exe` file.
   - Click the **Download** button or **View raw**, then save the file.

### Step 2: Set Up the Folder on Your Desktop
1. **Create a Folder for the Tool**
   - Right-click on your **Desktop**, select **New > Folder**, and name it `CustomerOrdersTool`.

2. **Move `script.exe` to the Folder**
   - Open **File Explorer** (`Win + E`).
   - Go to your **Downloads** folder.
   - Drag and drop `script.exe` into the 'CustomerOrdersTool' folder we just created.

3. **Prepare Your Input File**
   - Place your Excel file (e.g., `UnfilteredOrders.xlsx`) in the same `CustomerOrdersTool` folder.
  
4. **IMPORTANT NOTE**
   - The unfiltered file must have no blank rows near the top. This will require some manual work.
   - Open the unfiltered excel file and cut and paste all data entries so that the header columns are in "Row 1"
   - It should look like this:
     ![formatted excel update](https://github.com/user-attachments/assets/a4e916b6-641b-4a60-9fae-3dc11debe5b1)



---

## Running `script.exe` from the Command Line

### Step 1: Open Command Prompt
- Press `Win + R`, type `cmd`, and press **Enter**. Or you can look up "Command Prompt" in your Windows search bar.

### Step 2: Navigate to the Folder Using "Copy as Path"
1. On your Desktop, **Right-click** the `CustomerOrdersTool` folder and select **Copy as path**
2. In Command Prompt, type `cd `, then **press ctrl V** to paste the copied path.
3. Press **Enter** to navigate to the folder.

   **Example:**
   ```sh
   cd "C:\Users\YourUsername\Desktop\CustomerOrdersTool"


## Step 3: Run the Script with Automatic Naming

### Use the following command, replacing the start and end dates with the target month:

```sh
script.exe "UnfilteredOrders.xlsx" "FilteredOrders_January.csv" "2025-01-01" "2025-01-31"
```

### Explanation of Inputs:
"UnfilteredOrders.xlsx": This is the input file, which should be the Excel file containing the customer orders data. Make sure the file is in the same folder as the script.exe or provide the full path to the file.

"FilteredOrders_January.csv": This is the output file. The script will save the filtered data into this CSV file. The filename should reflect the month you are filtering, such as "FilteredOrders_January.csv" for orders in January.

"2025-01-01": This is the start date for the filtering. The script will include all orders starting from this date (format: YYYY-MM-DD).

"2025-01-31": This is the end date for the filtering. The script will include all orders up to this date (format: YYYY-MM-DD).

**For different months, change the date range and output filename accordingly:**

```sh
script.exe "UnfilteredOrders.xlsx" "FilteredOrders_February.csv" "2025-02-01" "2025-02-28"
```

After the script runs, there should be a new csv file located in the same folder as the script.exe and unfilteredorders.xlsx file. This file will be filtered based off of the date range, and is parsing for the word in the description *"Flexible"*

### **IMPORTANT NOTE:**
If there is a typo in the description of the word "Flexible," this result *will not* be included in the filtered list -- hence, it is important to check for typos in the description to ensure there is no loss of data

