# Building the master workbook

This master workbook is the report that we create by getting data from the sources, loading them to the data model and then creating the report based on the loaded data.

---

## Step 1: The driver cell

The queries read a named cell, so it has to exist before they can refresh.

1. In a new workbook, rename a sheet to **`Control`**.
2. In cell **A1**, type a label like `Access key (set by the tool) ▶`. This is only for future maintenance or modification by us.
3. Click on the Cell **B1** to make it the active cell, leave it empty, then click the **Name Box** (left of the formula bar), type **`pAccessKey`**, and press Enter. That name is how the query finds this cell (the tool overwrites it per recipient).

Leave `Control` visible for now, we'll hide it at the end.

## Step 2: The SourceFolder parameter

1. **Data → Get Data → Launch Power Query Editor**.
2. **Home → Manage Parameters → New Parameter**.

   Name it **`SourceFolder`**,
   Type **Text**,
   Current Value = the full path to your `sample-data` folder,
   e.g. `C:\Tools\rls-report-distributor\sample-data`

## Step 3: Paste the five queries

For each file in the `queries` folder:
**Home → New Source → Other Sources → Blank Query**, then **Home → Advanced
Editor**, replace everything with the file's contents, click Done, and **rename the query** (right pane) to match the name in the file's header. Create them in this order so each one's references already exist:

| Query          | Load destination       |
| -------------- | ----------------------- |
| `AccessKey`    | Connection only        |
| `AllowedIds`   | Connection only        |
| `Unrestricted` | Connection only        |
| `Employees`    | **Load to Data Model** |
| `Sales`        | **Load to Data Model** |

"Connection only" = in **Home → Close & Load To…** choose **Only Create
Connection** (do **not** tick Add to Data Model).

After everything has been loaded as connection only, head to the Data tab, click Queries and Connections, right click on `Employees` and `Sales`, select **Load To…** and tick **Add this data to the Data Model**.

## Step 4: The relationship

1. **Data → Manage Data Model** (or the Power Pivot tab → Manage) to open the Power Pivot window.
2. Switch to **Diagram View** and drag **`Sales[EmployeeID]`** onto
   **`Employees[EmployeeID]`** to create a one-to-many relationship between Sales and Employees.

## Step 5: Build the report

On a normal sheet (name it `Report`): **Insert → PivotTable → From Data Model**. Build one or two pivots off the model, for example:

- Rows = `Employees[FullName]`, Values = `Sales[Amount]`.
- Rows = `The whole Employees Table` to see which rows are present.

Everything here will show only the recipient's permitted data once the tool sets their access key, because the model itself only ever contains their rows.

## Step 6: Test it manually the first time

Before automating anything, we need to prove the security works manually:

1. On `Control`, set **B1** (`pAccessKey`) to **`3`**. **Data → Refresh All**.
   The pivots now show only Michael Blythe's data. Open the Power Pivot window and look at the `Sales` table: it contains only EmployeeID 3's rows. Nobody else's data is anywhere in the file.
2. Set `pAccessKey` to **`3,4`** and Refresh All. You'll see reps 3 and 4 (a manager's view), and rep 5 is absent everywhere.

If those refreshes behave, the row-level security is correct and the tool
just automates exactly this: write the key, refresh, save a copy, deliver.

## Step 7: Finish

1. Right-click the **`Control`** sheet tab → **Hide**.
2. **Save** the workbook as **`Sales-RLS-Master.xlsx`** somewhere, e.g.
   `C:\Tools\rls-report-distributor\master-reports`.
3. Add the master's folder as an Excel **Trusted Location** 
    (File → Options → Trust Center → Trust Center Settings → Trusted Locations).

We point `config.toml` at this file. From here on, the tool takes over.

The master workbook is done. Continue with the steps in the main README.
