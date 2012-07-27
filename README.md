MailMerge Done Right
====================
Google Apps Mail Merge script.
Author: Alex Skrypnyk (alex.designworks@gmail.com)
License: GPL v2+ (http://www.gnu.org/licenses/old-licenses/gpl-2.0.html)

Google Apps mail Merge script that uses Gmail Drafts as a source of template.
Allows using attachments and inline images.
Also supports contact information importing from Contacts groups.

1. Write your email and save it as Draft.
   You may attach images and attachments as you would normally do.
2. The words that needs to be replaced for each recipient are called
   "placeholders".  Replace all placeholders with unique names, surrounded
   by %% and %%.  You may use any characters, including spaces, in your
   placeholder name.
   Example: Hello %%First Name%%.
3. Create a separate column for each placeholder in the spreadsheet and fill
   it with values.
   First cell of each column (column header) must be the same as you have
   specified in the body of your email.
   Example:  for placeholder %%First Name%%, column name will be First Name
4. Run mail merge.

To re-send the message to selected recipients, clear cell in 'Sent status'
column for this recipient.

It is a good practice to create separate spreadsheet for each mail merge to
be able to track all sent correspondence.

Installation
============
You may try to find this script in Google Templates Gallery and copy to your
account.
OR
You may copy/paste this code directly into new spreadsheet in your Google Drive
using Script Editor (Tools->Script Editor) and save the script.
Refresh the spreadsheet and  'Mail Merge' menu item will appear after
5-7 seconds.

Video Tutorial
==============
@see http://www.youtube.com/watch?v=WWb3hpXLrag

Project Page
============
@see https://github.com/alexdesignworks/MailMerge-Done-Right
