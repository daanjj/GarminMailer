# GarminMailer

Simple app to mount Garmin watches, copy the most recent file(s) to SSD and then mail them to the email address given by the user. To be used in running workshops.

## Setup

This application can copy files from a watch without any configuration.

To send emails, you must create a configuration file named `mailer.conf.json`. This file should be placed in a `GarminMailer` folder inside your user's `Documents` directory.

*   **On macOS**, the path is: `~/Documents/GarminMailer/mailer.conf.json`
*   **On Windows**, the path is: `C:\Users\<YourUser>\Documents\GarminMailer\mailer.conf.json`

### Steps to Configure Email

1.  Create the `GarminMailer` folder inside your `Documents` folder if it doesn't already exist.
2.  Copy the `mailer.conf.json.example` file from the application directory into the `GarminMailer` folder you just created.
3.  Rename the copied file to `mailer.conf.json`.
2.  Edit the new file and enter your SMTP server details.

For Gmail, you will need to create an App Password and use that in the `password` field.
