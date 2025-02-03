# Google Sheets Membership Points Tracker

This script automates tracking and updating membership points for users based on Google Form submissions. It retrieves form data, updates the main tracking sheet, and ensures duplicate entries do not add extra points. Additionally, a car ride feature can be added for bonus points.

## Features

- **Automatic Data Retrieval**: Extracts form responses from the last submitted row.
- **Points Tracking**: Updates the main sheet with points for each participant.
- **New User Handling**: Adds a new row if a user is not already listed.
- **Duplicate Prevention**: Ensures multiple submissions for the same event do not give extra points.
- **Name Normalization**: Converts names to lowercase and trims whitespace for consistency.

## How It Works

1. A Google Form is created for each event.
2. The script extracts the latest form response from the Google Sheet.
3. It checks if the participant is already in the main sheet:
   - If found, their points are incremented.
   - If not, a new row is created with their details.
4. Name standardization prevents case-sensitive duplicates.
5. A separate script may be needed to track users’ event attendance over time.

## Setup Instructions

1. **Create a Google Form** for an event.
2. **Link the Form to a Google Sheet** (Responses are automatically recorded).
![Link Event for Each Google Form](images/linkeventforeachgoogleform.png)
    **keep track of the name of the sheet corresponding to the google form**
3. **Copy and paste the script** into the Google Apps Script Editor.
4. **Modify the following variables in the script**:
![After Linking Google Form with Google Sheet, Use the Name of the tab you have Created](images/afterlinkusethisname.png)
   - `sourceSheet`: The name of the sheet containing form responses.
    **this is the name of the sheet for the corresponding linked google form** 
   - `eventName`: The name of the event (used to track attendance).
        **this is what is shown up on the main google sheet** 
5. **set up a trigger on google appscript**
    a. go into triggers 
    b. add trigger 
        > should say 
            - onFormSubmit(or function name if you rename)
            - Head
            - From form
            - on form submit
![How Triggers Should Look](images/howtriggersshouldlook.png)






## Other Notes
**may have to rename if the sheet you want to put the data onto is not called 'main'** 