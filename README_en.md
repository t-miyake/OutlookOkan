Outlook Okan (Add-in to prevent mis-sent emails.)
========

日本語は[こちら](https://github.com/t-miyake/OutlookOkan/)。

Outlook Okan is an add-in for Microsoft Office Outlook.  

This add-in will display a confirmation window before sending an email.  
That's to prevent mis-sent emails.  

For sensitive emails, you can rest assured that this add-in is completely open source.  
There are also useful optional features such as keyword warnings and automatic Cc/Bcc addition.  

You can download this add-in [here](https://github.com/t-miyake/OutlookOkan/releases).  

This add-in is open source and free to use, but it is unsupported and unguaranteed.  
([License](https://github.com/t-miyake/OutlookOkan/blob/master/LICENSE))  
If you need customization or support, please contact us on an individual basis.  

Confirmation window before sending.  
![Screenshot 1](https://github.com/t-miyake/OutlookOkan/blob/master/Screenshots/en/Screenshot_v2.5.0_01_en.png)  

Settings window (general settings)  
![Screenshot 2](https://github.com/t-miyake/OutlookOkan/blob/master/Screenshots/en/Screenshot_v2.5.0_04_en.png)  

Settings window (deferred delivery)  
![Screenshot 3](https://github.com/t-miyake/OutlookOkan/blob/master/Screenshots/en/Screenshot_v2.5.0_05_en.png)  

Alert window  
![Screenshot 4](https://github.com/t-miyake/OutlookOkan/blob/master/Screenshots/en/Screenshot_v2.5.0_03_en.png)  

About window  
![Screenshot 5](https://github.com/t-miyake/OutlookOkan/blob/master/Screenshots/en/Screenshot_v2.5.0_02_en.png)  

## System requirements

- Windows 7 / 8 / 8.1 / 10
- Microsoft Outlook 2013 / 2016 / 2019 / Microsoft 365 Apps (32bit or 64bit)
- .NET Framework 4.6.2 or above

## Function list (overview)

- Confirmation before sending an email.  
  - Confirmation window is displayed before sending a mail and all items must be checked before sending a mail.
  - It is also possible not to display the confirmation before sending, such as an email to the internal domain.
  - External domains are shown in red letters.
  - Show subject and sender addresses, a list of attachments, and the body of the email.
  - Warn of missing attachments or large attachments.
  - Expand distribution lists and contact groups to show each recipient. (can be turned on or off)  
  - Sort and display the recipients by domain. (can be turned on or off)  

- Prohibit the sending of mails that match the conditions.
  - Prohibit the sending of emails to the specified destination or domain.
  - Prohibit the sending of emails containing the specified keyword in the body.
  - Prohibit the sending of mail when the number of destinations (To/Cc) external domains is large.

- Allowlist
  - Allowlisted domains and email addresses do not need to be checked on the confirmation winodw.

- Name and recipient registration and alerts
  - If the name in the body of the message and the address or domain of the recipient do not match, a alertings is displayed.

- Registering alerting keywords and alerting messages.
  - If the registered keyword is included in the body of an email, the registered alerting message will be displayed.

- Registering alerting recipients and alerting messages.
  - A alerting message is displayed when sending an email to the registered address or domain.

- Warning on the number of external domains of the destination (To/Cc) and auto change to Bcc.
  - Warning message when the number of destination (To/Cc) external domains is large.
  - When the number of destination (To/Cc) external domains is large, the destination (To/Cc) external domains is automatically changed to Bcc.

- Automatic Cc/Bcc addition(by keywords)
  - If the specified keyword is included in the body of an email, the specified address is automatically added to Cc and Bcc.

- Automatic Cc/Bcc addition(by recipients)
  - Automatically add the specified address to a Cc or Bcc in an email to the specified recipients.

- Automatic Cc/Bcc addition(by attachment)
  - Automatically add the specified address to Cc and Bcc in emails with files attached.

- Deferred delivery(Delayed delivery)
  - You can delay (put on hold) the sending of an email for a set amount of time (in minutes).
  - You can set a default delay time for each domain or email address.

- Importing and exporting settings
  - You can import and export your settings as a CSV file.

- Multi-language support
  - Both Japanese and English are available. It is also designed to allow for the addition of languages.

## Manual

[Wiki(Manual)](https://github.com/t-miyake/OutlookOkan/wiki/Manual)  

## Known Issues

[Wiki(Known Issues)](https://github.com/t-miyake/OutlookOkan/wiki/Known-Issues)  

## Roadmap

[Wiki(Roadmap)](https://github.com/t-miyake/OutlookOkan/wiki/Roadmap)  
