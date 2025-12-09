import argparse

from .mailer import Mailer


def build_parser():
    parser = argparse.ArgumentParser(description='Send emails via Outlook using winoutlook.')

    sub = parser.add_subparsers(dest='command')

    send = sub.add_parser('send', help='Send an email')
    send.add_argument('--to', required=True)
    send.add_argument('--subject', required=True)
    send.add_argument('--body', default='')
    send.add_argument('--html', action='store_true')
    send.add_argument('--attach', nargs='*', default=[])
    send.add_argument('--use-account', default=None)
    send.add_argument('--use-name', default=None)

    return parser


def main():
    parser = build_parser()
    args = parser.parse_args()

    if args.command == 'send':
        mailer = Mailer()
        mailer.send_email(
            recipient=args.to,
            subject=args.subject,
            body=args.body,
            html_format=args.html,
            attachments=args.attach,
            send_using_account=args.use_account,
            sent_on_behalf_of_name=args.use_name,
        )
