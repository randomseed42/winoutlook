import argparse

from .mailer import Mailer


def build_parser():
    parser = argparse.ArgumentParser(description='Send emails via Outlook using winoutlook.')

    parser.add_argument('-t', '--to', required=True, help='Recipient email address')
    parser.add_argument('-s', '--subject', required=True, help='Email subject')
    parser.add_argument('-b', '--body', default='', help='Email body text (default is empty)')
    parser.add_argument('-f', '--html', action='store_true', help='Send email as HTML (default is plain text)')
    parser.add_argument(
        '-a',
        '--attach',
        nargs='*',
        default=[],
        help='Attach files to the email (e.g., -a file1.pdf file2.docx)',
    )
    parser.add_argument(
        '-u',
        '--use-account',
        default=None,
        help='Specify the Outlook account to use (default is the default account)',
    )
    parser.add_argument(
        '-n',
        '--use-name',
        default=None,
        help='Specify the sender name (default is the account name)',
    )

    return parser


def main():
    parser = build_parser()
    args = parser.parse_args()

    Mailer.send_email(
        recipient=args.to,
        subject=args.subject,
        body=args.body,
        html_format=args.html,
        attachments=args.attach,
        send_using_account=args.use_account,
        sent_on_behalf_of_name=args.use_name,
    )
