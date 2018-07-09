# Y Drive Accelerator
This is an unpublished, developer Gmail add-on for quickly saving emails to a team drive

## Background
In the Office of Research, the Sponsored Projects team has to save copies of almost all of their email correspondence. They communicate with funding agencies, other universities, other departments at UCSB, etc. so they can receive hundreds of emails every day. All of these emails need to be saved to a specific location on our network, a shared drive called the Y Drive. This drive is organized by fiscal year, department, and some other categories. The process of manually converting an email to a pdf, finding the right location to save it to, and renaming it is a lengthy process. When you get as many emails as they do, this can end up taking a significant amount of time every day.

## The Solution
We decided to create a Gmail add-on, the Y Drive Accelerator, to fix this problem. When our users select the email they want to save, they can choose the file name, whether to include all emails in the thread or just the selected one, and whether or not to include the email's attachments. They also enter a record number that is used to determine the right location to save the file(s). Once they enter this information and click 'Save', the add-on creates a temporary html file from the emails and saves it to a Google team drive as a pdf (along with the attachments if they chose to include them). We have a Windows Service that constantly checks the team drive and downloads the files to the correct location in our file system.

## Installation
** NOTE: This add-on will only work if you have upgraded to the "New" version of Gmail **
* Create a new Google Apps Script project (Drive->New->More->Google Apps Script)
* Copy the files from this repository into the apps script project
* Create a team drive in Google Drive with a folder for the add-on to save to
* Rename var _ROOTFOLDER at the top of Code.gs to the name of the folder in the team drive
* In Apps Script, go to Publish->Deploy from manifest->Get ID
* Copy the deployment ID
* In Gmail, go to Settings->Add-ons and check the box labeled "Enable developer add-ons for my account"
* Enter the deployment ID in the developer add-ons textbox and click "Install"
* Refresh Gmail and select an email. You should see the add-on icon on the right side of the screen

## Usage
* Select an email and click the add-on icon
* For the record number/subaward number textbox, enter a number in the following format 20######
* Enter a new file name, or use the prepopulated name (email subject)
* Choose to save the thread from the selected email, save the thread from the most recent email, or save just the selected email
* Choose whether or not to include attachments
(Note that these options always display, even if the email is not part of a thread or if there are no attachments)
* Click "Save". You should see a notification that says "Thread Saved"
* Open the folder in the team drive. You should see a subfolder named 20###### that has a pdf of the email, any attachments, and 2 other files

## Notes
* If you receive an error saying "Rootfolder not found", make sure you created a folder in the team drive and open the folder. This is a one time step
* The add-on works on mobile devices. If you install it on your computer and you have a new version of Gmail installed on your mobile device, it should show up there without any extra steps
* I included an outline for how the Windows Service works. I won't describe this here, but please let me know if you'd like more information

## Limitations
* The biggest limitation of the add-on is due to a Google quota. Custom functions are only allowed to run for 30 seconds. If they run longer than that, the add-on will stop working. Emails with many attachments will hit this limit. Each save to the team drive takes anywhere from 1-3 seconds, which adds up quickly. We don't have a solution in place for this yet as it doesn't occur very often for us. I have a couple ideas for this problem that wouldn't be perfect so I'm open to suggestions!
* The Card UI that the add-on uses is still very new, so there are some things that you would expect to be available that aren't. Nothing too important is missing though
* Because it is a developer add-on, users have to have read access to the actual script. You have to share the Apps Script project with every new user, and also assign them Edit access in the team drive

## Author
Cameron McNair - Applications Developer - cfb@ucsb.edu

Please reach out to me if you have any questions or comments. I liked this project and I'm happy to talk about it!
