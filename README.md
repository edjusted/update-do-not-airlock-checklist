# update-do-not-airlock-checklist

# Introduction
If you play Star Trek Timelines, there is a Google Sheets (aka GS) tool available called the [Do Not Airlock checklist](https://forum.disruptorbeam.com/stt/discussion/15561/do-not-airlock-checklist-thread-4/p1).

If you use that Google Sheet, these scripts make it easy to import/copy over your info from your old GS to the new one when there's an update.

As of v5.0.1, the Do Not Airlock checklist GS already has this functionality built in, so the Installation section is not necessary.

# Installation

1. Make a copy of the new Checklist and create a new tab called **UpgradeFromOldGs**.

2. Go to the **Extensions** menu and select **Apps Script**.\
A new tab should open and you'll probably see some code there already, and it probably starts with `function pullData() {`\
You'll want to ignore it and leave it alone. 

3. Copy the contents from the [Code.gs](https://github.com/edjusted/update-do-not-airlock-checklist/blob/main/Code.gs) file here and paste it into the Script Editor tab *below* any existing code (i.e. don't overwrite any existing code).\
Click the save button to save.

4. Go back to the GS window and reload the page.

# Usage

0. First, go to your old Google Sheet (aka GS) and copy the url.

1. Come back to the new GS, go to the **UpgradeFromOldGs** tab. From here on, everything you do will be on this tab.

2. Paste the url of your *old* Google Sheet into cell B1. Note that it *must* be in cell B1.

3. If you *do* use Datacore (i.e. you go to Datacore and export your info, then import it into the **Import** tab),\
   there is no step 3. Leave the checkbox in B2 unchecked.

   If you *don't* use Datacore (i.e. you manually fill in the "fused", "Level", and "active" columns for your crew on the **Crew** tab),\
   check the checkbox in B2. This will copy over all of your manually typed-in info.

4. Click on the **Upgrade from old Checklist** menu, and select one of the following:
- Import/overwrite Crew info (Fused / Level / Active / Keep / Cite / Notes cols)
- Import/overwrite Missions
- Import/overwrite Settings
- Copy over user tabs (see Options below)

Each function need to be called separately because this GS is extremely formula-intensive, and trying to copy over too much at once can cause timeout issues.
(Some of the functions may take upwards of 2-3 mins. Please be patient.)\
The best thing to do is select each one, wait for the GS to finish processing, then select the next one. Rinse & repeat.

(You'll actually need to do this **twice** the *first* time. Keep reading.)\
    **Important**: The first time you select this option, Google will ask you to grant permissions to the script. Go through the steps and **Allow** the permissions.\
          ![Google Authorization](.github/google-authorization-screenshot.png)\
    You will then need to select the menu option *again* to actually run it. (The first "run" kicks off the permission process but doesn't actually run the scripts.)

  **Import/overwrite Crew info (Fused / Level / Active / Keep / Cite / Notes cols)**
  **What this does**: This copies over your **Crew** tab info (i.e. the info in the Fused, Level, Active, Cite, and Notes columns)
If the checkbox in B2 is *unchecked*, this will also copy over the **Import** tab.

  **Import/overwrite Missions**
  **What this does**: This copies over your settings/selections on the **Missions** tab

  **Import/overwrite Settings**
  **What this does**: This copies over your settings/selections on the **Settings** tab

  **Copy over custom tabs (see Options below)**
  **What this does**: If you've created any new/additional tabs in the old GS for your own use, and you want to copy them over, use this menu.
But first...
...in cell B3, type in the exact name(s) of the tab(s) you want to copy over, separated by a comma if there are multiple tabs.


# Troubleshooting

If any of the imports time out, just wait for the GS to finish processing anything, then try it again. Usually that will fix the problem.

Or if you still have timeout issues after trying it more than once, you can also try to run the "sub-processes" separately.\

Go to the **Upgrade from old Checklist** menu, then the **For Troubleshooting** sub-menu, and select the following:\
    • **Import/overwrite Crew (manual Fused / Level / Active cols)** - this option is for people who *don't* use Datacore. This copies over the *values* from the "Fused", "Level", and "Active" columns on your **Crew** tab.\
    • **Import/overwrite Crew (Import tab)** - this option is for people who *do* use Datacore. This copies over the **Import** tab.\
Or you can just go to datacore and re-download / re-import a fresh batch of info.\
    • **Import/overwrite Crew Notes/Keep/Cite cols** - this copies over the "Keep", "Cite", and "Notes" columns from your Crew tab.
