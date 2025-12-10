# winoutlook

[![PyPI version](https://badge.fury.io/py/winoutlook.svg)](https://badge.fury.io/py/winoutlook)
[![Python version](https://img.shields.io/badge/python-3.13+-blue.svg)](https://www.python.org/downloads/release/python-3130/)
[![License](https://img.shields.io/badge/license-Apache%202.0-blue.svg)](https://github.com/randomseed42/winoutlook/blob/main/LICENSE)


## Overview

`winoutlook` is a Windows-only Python library for sending emails through locally installed Microsoft Outlook using the win32 COM interface. It provides a simple command-line interface (CLI) to send emails directly from your Python scripts or command line.

## Features

- **Send Emails**: Easily send emails with plain text or HTML content.
- **Attachments**: Attach multiple files to your emails.
- **Multiple Accounts**: Specify which Outlook account to use for sending emails.
- **Sender Name**: Send emails on behalf of a different sender name.

## Installation

### Using pip

You can install `winoutlook` via pip:

```bash
pip install winoutlook
```

### Requirements

- Python 3.13 or later
- Microsoft Outlook installed on your Windows machine

## Usage

### Command-Line Interface (CLI)

The `winoutlook` library comes with a CLI tool called `wosend` that you can use to send emails directly from the command line.

```bash
wosend -t recipient@example.com -s "Subject Here" -b "This is the body of the email."
```

#### Options

- `-t, --to`: Recipient email address (required).
- `-s, --subject`: Email subject (required).
- `-b, --body`: Email body text (default is empty).
- `-f, --html`: Send email as HTML (default is plain text).
- `-a, --attach`: Attach files to the email (e.g., `-a file1.pdf file2.docx`).
- `-u, --use-account`: Specify the Outlook account to use (default is the default account).
- `-n, --use-name`: Specify the sender name (default is the account name).

### Python API

You can also use `winoutlook` directly in your Python scripts.

```python
from winoutlook.mailer import Mailer

Mailer.send_email(
    recipient='recipient@example.com',
    subject='Subject Here',
    body='This is the body of the email.',
    html_format=False,
    attachments=['file1.pdf', 'file2.docx'],
    send_using_account='youraccount@example.com',
    sent_on_behalf_of_name='Your Name',
)
```

#### Parameters

- `recipient`: Recipient email address (required).
- `subject`: Email subject (required).
- `body`: Email body text (default is empty).
- `html_format`: Send email as HTML (default is plain text).
- `attachments`: List of file paths to attach (default is none).
- `send_using_account`: Specify the Outlook account to use (default is the default account).
- `sent_on_behalf_of_name`: Specify the sender name (default is the account name).

## Examples

### Sending a Plain Text Email

```bash
wosend -t recipient@example.com -s "Hello" -b "This is a test email."
```

### Sending an HTML Email

```bash
wosend -t recipient@example.com -s "Hello" -b "<h1>This is a test email.</h1>" -f
```

### Sending an Email with Attachments

```bash
wosend -t recipient@example.com -s "Hello" -b "Please find the attached files." -a file1.pdf file2.docx
```

### Sending an Email Using a Specific Account

```bash
wosend -t recipient@example.com -s "Hello" -b "This is a test email." -u youraccount@example.com
```

### Sending an Email on Behalf of Someone Else

```bash
wosend -t recipient@example.com -s "Hello" -b "This is a test email." -n "Your Name"
```


## License

This project is licensed under the Apache License 2.0. See the [LICENSE](LICENSE) file for details.

## Contact

For any questions or support, please contact the maintainer:

- **Name**: randomseed42
- **Email**: randomseed42@protonmail.com

## Acknowledgments

- Thanks to the [pywin32](https://github.com/mhammond/pywin32) project for providing the necessary tools to interact with the win32 COM interface.
- Special thanks to the [email-validator](https://github.com/JoshData/python-email-validator) library for email validation.

## Changelog

- **v0.1.0**: Initial release with basic email sending functionality.
