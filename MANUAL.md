# GarminMailer User Manual

Welcome to GarminMailer! This guide will walk you through setting up and using the application on your macOS or Windows computer.

## Table of Contents

1.  [What is GarminMailer?](#what-is-garminmailer)
2.  [Installation](#installation)
    *   [macOS](#macos)
    *   [Windows](#windows)
3.  [Modes of Operation](#modes-of-operation)
4.  [Configuration (for Email Mode)](#configuration-for-email-mode)
    *   [Creating the Configuration File](#creating-the-configuration-file)
    *   [Using Gmail and App Passwords](#using-gmail-and-app-passwords)
5.  [How to Use GarminMailer](#how-to-use-garminmailer)
    *   [Email Mode](#email-mode)
    *   [Copy-Only Mode](#copy-only-mode)
6.  [Troubleshooting](#troubleshooting)

---

## What is GarminMailer?

GarminMailer is a simple desktop application designed for workshops and events. It allows users to quickly copy activity files (.fit files) from a Garmin watch to the computer and, optionally, email them to a specified address.

It has two main modes:
*   **Email Mode**: Copies activity files and sends them as attachments in an email.
*   **Copy-Only Mode**: Copies activity files to a local folder without sending an email.

## Installation

First, download the latest version of the application from the [GitHub Releases page](https://github.com/DaanDeMeyer/GarminMailer/releases).

### macOS

1.  Download the `GarminMailer-macOS.zip` file.
2.  Unzip the file. This will create a `GarminMailer.app` file.
3.  Move `GarminMailer.app` to your `Applications` folder.
4.  The first time you open the app, you may see a security warning. Right-click the app icon and select **Open**. You will only need to do this once.

### Windows

1.  Download the `GarminMailer.exe` file.
2.  Place the `.exe` file in a convenient location, such as your Desktop or a dedicated folder.
3.  Double-click `GarminMailer.exe` to run it. You may see a Windows SmartScreen prompt; click "More info" and then "Run anyway".

## Modes of Operation

The application's behavior is controlled by the **"Copy only, do not send mail"** checkbox:

*   **Unchecked (Default)**: The application is in **Email Mode**. It requires a name, a recipient email, and a one-time configuration to send emails.
*   **Checked**: The application is in **Copy-Only Mode**. It does not require any configuration and will automatically start when a watch is connected.

## Configuration (for Email Mode)

To send emails, you must create a configuration file containing your SMTP server details. This is a one-time setup.

### Creating the Configuration File

1.  **Create a folder**: In your user's `Documents` folder, create a new folder named `GarminMailer`.
    *   **macOS Path**: `~/Documents/GarminMailer/`
    *   **Windows Path**: `C:\Users\<YourUser>\Documents\GarminMailer\`

2.  **Create the config file**: Inside the `GarminMailer` folder, create a new text file named `mailer.conf.json`.

3.  **Add your credentials**: Open `mailer.conf.json` with a text editor (like Notepad or TextEdit) and paste the following content. Replace the placeholder values with your actual email credentials.

    ```json
    {
      "smtp_server": "smtp.gmail.com",
      "smtp_port": 465,
      "username": "your-email@gmail.com",
      "password": "your-google-app-password"
    }
    ```

### Using Gmail and App Passwords

If you use Gmail, you cannot use your regular Google password. You must generate an **App Password**.

1.  Go to your Google Account settings: myaccount.google.com.
2.  Navigate to **Security**.
3.  Under "How you sign in to Google," ensure **2-Step Verification** is turned **On**.
4.  Click on **App passwords**.
5.  Generate a new password:
    *   For "Select app," choose **Mail**.
    *   For "Select device," choose **Other (Custom name)** and type "GarminMailer".
    *   Click **Generate**.
6.  Copy the 16-character password that appears and paste it into the `"password"` field in your `mailer.conf.json` file.

## How to Use GarminMailer

### Email Mode

Use this mode to send activity files to participants.

1.  Launch the GarminMailer application.
2.  Ensure the "Copy only..." checkbox is **unchecked**.
3.  Enter the user's **Name** and **Recipient email**.
4.  Connect a Garmin watch to the computer via USB.
5.  Click the **Submit** button.
6.  The application will detect the watch, copy the latest activity file(s), and send the email.
    *   If multiple activities from today are found, a dialog will appear asking you to select which ones to send.
7.  After a successful send, the Name and Email fields will clear, ready for the next user.

### Copy-Only Mode

Use this mode to quickly back up files from multiple watches to your computer.

1.  Launch the GarminMailer application.
2.  Check the **"Copy only, do not send mail"** checkbox. The Name and Email fields will be disabled.
3.  Connect a Garmin watch. The process will start automatically.
4.  A dialog will appear showing the 5 most recent activities. Select the file(s) you wish to copy and click **Select**.
5.  The files are copied to a dated subfolder within `Documents/GarminMailer/sent/`.
6.  The watch will be ejected, and a message will appear prompting you to connect the next watch.

## Troubleshooting

*   **"No Garmin watch detected"**: Ensure the watch is connected properly and is in "Mass Storage" or "File Transfer" mode (not "Garmin" mode). Try a different USB port or cable.
*   **"Config not found"**: Make sure the `mailer.conf.json` file is correctly named and located in the `Documents/GarminMailer` folder.
*   **"AUTH: Gmail rejected login"**: This almost always means you are using your regular password instead of an **App Password**. Please see the configuration section above.
*   **macOS Security Warning**: If the app won't open, right-click the app icon and select **Open**.
*   **Files are not found**: Make sure the watch has `.fit` activity files stored in its `GARMIN/Activity/` folder.

---