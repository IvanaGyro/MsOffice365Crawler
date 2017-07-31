#-*- coding:utf8 -*-
import OfficeUsed
import textwrap
import argparse
import sys

parser = argparse.ArgumentParser(formatter_class=argparse.RawDescriptionHelpFormatter, description='''\
Auto crawl the number of computers that have been registerated in the specific MS office account.

actions:
    login-crawl
    others\
''')

parser.add_argument('action', help='Use username and password to login and crawl the numbers.')

parser.add_argument('-a', '--all', action='store_true', help='Crawl all accounts in the account file.')
parser.add_argument('-f', '--account-file', nargs=1, default=[OfficeUsed.defaultAccountFn], metavar='filename', help='File path of the account list.')
parser.add_argument('-o', '--output-file', nargs=1, default=[OfficeUsed.defaultInstalledFn], metavar='filename', help='File path of the output file.')
args = parser.parse_args()

print(args)

if args.action == 'login-crawl':
    if all:
        OfficeUsed.save_all_installsUsedVal_with_login(accountFn=args.account_file[0], outputFn=args.output_file[0])
else:
    print('Wrong action!')