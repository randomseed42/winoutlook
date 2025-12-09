import os
import sys
from pathlib import Path
from time import sleep, time

import pythoncom
import win32com.client as win32
from email_validator import EmailNotValidError, validate_email
from win32com.client.dynamic import CDispatch


class Mailer:
    @staticmethod
    def email_normalizer(email: str) -> str:
        """Normalize email address by removing any whitespace and converting to lowercase."""
        try:
            email = validate_email(email.strip()).normalized.lower()
            return email
        except EmailNotValidError as err:
            print(f'Invalid email address {email}: {err}')
            sys.exit(1)
        except Exception as err:
            raise Exception(f'Error normalizing email address {email}: {err}') from err

    @staticmethod
    def _get_outlook() -> CDispatch:
        """Return the Outlook application object."""
        timeout = 20
        # pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        pythoncom.CoInitialize()
        outlook = win32.Dispatch('Outlook.Application')
        # try:
        #     outlook = win32.GetActiveObject('Outlook.Application')
        # except pythoncom.com_error:
        #     outlook = win32.Dispatch('Outlook.Application')

        start = time()
        while True:
            try:
                namespace = outlook.GetNamespace('MAPI')
                _ = namespace.Folders.Count
                return outlook
            except Exception as err:
                if time() - start > timeout:
                    pythoncom.CoUninitialize()
                    raise TimeoutError(f'Timeout waiting for Outlook: {err}') from err
                sleep(1)

    @classmethod
    def _get_default_account(cls) -> tuple[str, str]:
        """Return the default account for the current Outlook session."""
        outlook = cls._get_outlook()

        namespace = outlook.GetNamespace('MAPI')
        address_entry = namespace.CurrentUser.AddressEntry
        if address_entry.Type == 'EX':
            exchange_user = address_entry.GetExchangeUser()
            if not exchange_user:
                raise ValueError('Exchange account not found.')
            name = exchange_user.Name
            address = cls.email_normalizer(exchange_user.PrimarySmtpAddress)
            return name, address
        else:
            name = address_entry.Name
            address = cls.email_normalizer(address_entry.Address)
            return name, address


    @classmethod
    def get_default_account(cls) -> tuple[str, str]:
        """Return the default account for the current Outlook session."""
        try:
            name, address = cls._get_default_account()
            return name, address
        except Exception as err:
            raise Exception(f'Error getting default account: {err}') from err
        finally:
            pythoncom.CoUninitialize()

    @classmethod
    def _send_email(
        cls,
        recipient: str,
        subject: str,
        body: str,
        html_format: bool = False,
        attachments: str | os.PathLike | list[str | os.PathLike] | None = None,
        send_using_account: str | None = None,
        sent_on_behalf_of_name: str | None = None,
    ):
        """Send an email with optional attachments."""
        sender_name, sender_email = cls.get_default_account()
        recipient = cls.email_normalizer(recipient)
        send_using_account = send_using_account or sender_email
        send_using_account = cls.email_normalizer(send_using_account)
        sent_on_behalf_of_name = sent_on_behalf_of_name or sender_name

        outlook = cls._get_outlook()
        mail = outlook.CreateItem(0)

        mail.Sender = sender_email
        mail.SendUsingAccount = send_using_account
        mail.SentOnBehalfOfName = sent_on_behalf_of_name
        mail.To = recipient

        mail.Subject = subject

        if html_format:
            mail.BodyFormat = 2
            mail.HTMLBody = body
        else:
            mail.BodyFormat = 1
            mail.Body = body

        if isinstance(attachments, str | os.PathLike):
            if not Path(attachments).is_file():
                raise ValueError(f'File not found: {attachments}')
            attachments = [attachments]

        if attachments:
            _attachments = []
            for attachment in attachments:
                if not Path(attachment).is_file():
                    raise ValueError(f'File not found: {attachment}')
                if Path(attachment).stat().st_size > 10 * 1024 * 1024:
                    raise ValueError(f'File size exceeds 10MB: {attachment}')
                if Path(attachment).suffix.lower() not in ('.txt', '.pdf', '.docx', 'csv', 'xlsx', '.jpg', '.png'):
                    raise ValueError(f'Unsupported file type: {attachment}')
                if attachment in _attachments:
                    raise ValueError(f'Duplicate attachment: {attachment}')
                _attachments.append(attachment)
            for _attachment in _attachments:
                mail.Attachments.Add(_attachment)

        mail.Send()

    @classmethod
    def send_email(
        cls,
        recipient: str,
        subject: str,
        body: str,
        html_format: bool = False,
        attachments: str | os.PathLike | list[str | os.PathLike] | None = None,
        send_using_account: str | None = None,
        sent_on_behalf_of_name: str | None = None,
    ):
        """Send an email with optional attachments."""
        try:
            cls._send_email(
                recipient=recipient,
                subject=subject,
                body=body,
                html_format=html_format,
                attachments=attachments,
                send_using_account=send_using_account,
                sent_on_behalf_of_name=sent_on_behalf_of_name,
            )
        except Exception as err:
            raise Exception(f'Error sending email: {err}') from err
        finally:
            pythoncom.CoUninitialize()

    @classmethod
    def __send_email(
        cls,
        recipient: str,
        subject: str,
        body: str,
        html_format: bool = False,
        attachments: str | os.PathLike | list[str | os.PathLike] | None = None,
        send_using_account: str | None = None,
        sent_on_behalf_of_name: str | None = None,
    ):
        """Send an email with optional attachments."""
        try:
            cls.send_email(
                recipient=recipient,
                subject=subject,
                body=body,
                html_format=html_format,
                attachments=attachments,
                send_using_account=send_using_account,
                sent_on_behalf_of_name=sent_on_behalf_of_name,
            )
        except Exception as err:
            raise Exception(f'Error sending email: {err}') from err
        finally:
            pythoncom.CoUninitialize()
