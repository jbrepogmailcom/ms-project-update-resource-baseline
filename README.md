# MS Project Update Resource Baseline

MS Project does not have the ability to recalculate resource baselined work summaries. This code does it.

> **Note:** All the code in this project is designed for **Baseline1 Work**. If you use a different baseline, you will need to adjust the code accordingly.

## Description

When you save work as a baseline in MS Project, the following data is saved:
- **Time-phased data for assignments** (saved as the baseline).
- **Current summaries for assignments** over the entire time period.
- **Time-phased data for the entire resource**.
- **Summary for the entire assignment** (for all time).
- **Total resource baseline work** (summed over all time).

However, when you later replan and save tasks that contain these assignments, MS Project updates **only**:
- Task-related data, and
- Time-phased data of the assignment in the baseline.

**Not updated:**
- Vertical summaries (time-phased resource total baselined work)
- Horizontal summaries (sum of all baselined work per assignment)
- Total resource baseline work

This script goes through all resource assignments and all project days to recalculate missing summaries. It also deletes any orphaned assignments (from deleted tasks or resources) since they are invalid.

## Usage

### Enabling the Developer Tab in MS Project

1. **Click** on the **File** tab.
2. **Select** **Options**.
3. In the **Project Options** window, choose **Customize Ribbon**.
4. In the right pane, **check** the box for **Developer**.
5. **Click** **OK**.

### Step-by-Step Instructions

#### 1. Open the Visual Basic Editor

You can open the Macro Editor (Visual Basic Editor) using one of these methods:

- **Keyboard Shortcut:** Press `Alt + F11` on your keyboard.
- **Via Ribbon:** Click on the **Developer** tab and then click **Visual Basic**.

#### 2. Insert a New Module

1. In the Visual Basic Editor (VBE), go to the menu and click **Insert**.
2. Select **Module**.  
   A new module window will open where you can add your code.

#### 3. Paste Your Visual Basic Code

1. Make sure the new module is active.
2. Press `Ctrl + V` (or right-click and select **Paste**) to insert your code from the clipboard into the module.

#### 4. Save Your Project

- In the VBE, click **File** > **Save [YourProjectName]** to save your changes.
- Alternatively, return to Microsoft Project and save the project file.

#### 5. Run the Macro

There are several ways to run your macro:

**Option A: Run from the Visual Basic Editor**
- With your code window open, press `F5` to run the macro.
- Alternatively, click **Run** > **Run Sub/UserForm** from the menu bar.

**Option B: Run from Microsoft Project via the Developer Tab**
1. In Microsoft Project, click on the **Developer** tab.
2. Click on **Macros**.
3. In the Macro dialog box, select the macro you want to run.
4. Click **Run**.

**Option C: Run the Macro Using Alt + F8**
1. In Microsoft Project, press `Alt + F8` to open the Macros dialog box.
2. Select the macro you want to run.
3. Click **Run**.

## Additional Notes

- This script is designed specifically for **Baseline1 Work**.
- It recalculates missing summaries for all resource assignments and project days.
- If it detects orphaned assignments (from deleted tasks or resources), it will delete them since they are invalid.
- If you experience issues, verify your macro security settings under **File** > **Options** > **Trust Center** > **Trust Center Settings** > **Macro Settings**.

---

Feel free to modify or extend this guide to suit your needs!
