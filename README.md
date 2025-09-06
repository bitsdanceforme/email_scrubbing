# email_scrubbing
Scripts discussed at BTC 2025

# Caveats
Google does throttle the number of accesses that can be made against an email account
Get familiar with the google API
https://developers.google.com/apps-script/reference/properties/properties#getProperty(String)

There is a quota limit: https://developers.google.com/apps-script/guides/services/quotas


# How to use
This script is to be used in google AppsScript (https://script.google.com/home).

1. Download the script
2. Create new AppsScript project
3. Upload pullEmails.gs
4. Use the View> Show Project Properties to setup the variables in the script
Find email sent to targetEmail in the last 15 years
deliveredto:myemail@gmail.com after:2009/01/01 before:2025/01/01

# How to use the stop feature
Open the Apps Script editor.

Go to View > Show project properties > Script properties (or use the Apps Script API).

Add a property:

Key: stopNow

Value: true

The next time main() runs, it will detect the flag, write the report, clear progress, and stop.

After stopping, the flag is automatically cleared so you can restart fresh later.

