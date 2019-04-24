import os,sys
import boto3
import datetime
import logging
import pandas as pd
import argparse

parser = argparse.ArgumentParser()
parser.add_argument('--months', type=int, default=3)
args = parser.parse_args()


# For date
from dateutil.relativedelta import relativedelta

# For email
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate

# Global Variables
if os.environ.get('SES_SEND'):
    SES_SEND = os.environ.get('SES_SEND')
    SES_REGION = os.environ.get('SES_REGION')
    SES_FROM = os.environ.get('SES_FROM')
else:
    print("SES Variables are required to set in Environments")
    sys.exit()

print(SES_SEND)
print(SES_REGION)
print(SES_FROM)


class CostExplorer:

    def __init__(self, CurrentMonth=False):
        # Array of reports ready to be output to Excel.
        self.reports = []

        self.client = boto3.client('ce')

        self.end = datetime.date.today().replace(day=1)

        self.riend = datetime.date.today()

        if CurrentMonth:
            self.end = self.riend

        self.start = (datetime.date.today() - relativedelta(months=+12))\
            .replace(day=1)  # 1st day of month 12 months ago

        self.ristart = (datetime.date.today() - relativedelta(months=+11))\
            .replace(day=1)  # 1st day of month 11 months ago

        self.sixmonth = (datetime.date.today() - relativedelta(months=+6))\
            .replace(day=1)  # 1st day of month 6 months ago, so RI util has savings values

        try:
            self.accounts = self.getAccounts()
        except:
            logging.info("Getting Account names failed")
            self.accounts = {}

    def setStart(self, interval):
        self.start = (datetime.date.today() - relativedelta(months=+interval)).replace(day=1)

    def getAccounts(self):
        accounts = {}
        client = boto3.client('organizations', region_name='us-east-2')
        paginator = client.get_paginator('list_accounts')
        response_iterator = paginator.paginate()
        for response in response_iterator:
            for acc in response['Accounts']:
                accounts[acc['Id']] = acc
        return accounts


    def addReport(self, Name="Default", GroupBy=[{"Type": "DIMENSION", "Key": "SERVICE"}, ],
                  Style='Total', NoCredits=True, CreditsOnly=False, RefundOnly=False, UpfrontOnly=False,
                  IncSupport=False):

        type = 'chart'  # other option table
        results = []

        Filter = {
            "Not": {"Dimensions": {"Key": "RECORD_TYPE", "Values": ["Credit", "Refund", "Upfront", "Support"]}}}
        if IncSupport:  # If global set for including support, we dont exclude it
            Filter = {"Not": {"Dimensions": {"Key": "RECORD_TYPE", "Values": ["Credit", "Refund", "Upfront"]}}}

        if CreditsOnly:
            Filter = {"Dimensions": {"Key": "RECORD_TYPE", "Values": ["Credit", ]}}
        if RefundOnly:
            Filter = {"Dimensions": {"Key": "RECORD_TYPE", "Values": ["Refund", ]}}
        if UpfrontOnly:
            Filter = {"Dimensions": {"Key": "RECORD_TYPE", "Values": ["Upfront", ]}}


        response = self.client.get_cost_and_usage(
            TimePeriod={
                'Start': self.start.isoformat(),
                'End': self.end.isoformat()
            },
            Granularity='MONTHLY',
            Metrics=[
                'UnblendedCost',
            ],
            GroupBy=GroupBy,
            Filter=Filter
        )

        if response:
            results.extend(response['ResultsByTime'])

            while 'nextToken' in response:
                nextToken = response['nextToken']
                response = self.client.get_cost_and_usage(
                    TimePeriod={
                        'Start': self.start.isoformat(),
                        'End': self.end.isoformat()
                    },
                    Granularity='MONTHLY',
                    Metrics=[
                        'UnblendedCost',
                    ],
                    GroupBy=GroupBy,
                    NextPageToken=nextToken
                )
                results.extend(response['ResultsByTime'])
                if 'nextToken' in response:
                    nextToken = response['nextToken']
                else:
                    nextToken = False

        rows = []
        sort = ''
        for v in results:
            row = {'date': v['TimePeriod']['Start']}
            sort = v['TimePeriod']['Start']
            for i in v['Groups']:
                key = i['Keys'][0]
                if key in self.accounts:
                    key = self.accounts[key]['Email']
                row.update({key: float(i['Metrics']['UnblendedCost']['Amount'])})

            if not v['Groups']:
                row.update({'Total': float(v['Total']['UnblendedCost']['Amount'])})
            rows.append(row)

        df = pd.DataFrame(rows)
        df.set_index("date", inplace=True)
        df = df.fillna(0.0)

        df = df.T
        df = df.sort_values(sort, ascending=False)
        self.reports.append({'Name': Name, 'Data': df, 'Type': type})

    def generateExcel(self):
        # Create a Pandas Excel writer using XlsxWriter as the engine.
        # os.chdir('/tmp')
        writer = pd.ExcelWriter('cost_explorer_report.xlsx', engine='xlsxwriter')
        workbook = writer.book
        for report in self.reports:
            print(report['Name'], report['Type'])
            report['Data'].to_excel(writer, sheet_name=report['Name'])
            worksheet = writer.sheets[report['Name']]
            if report['Type'] == 'chart':

                # Create a chart object.
                chart = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})

                chartend = 13
                for row_num in range(1, len(report['Data']) + 1):
                    chart.add_series({
                        'name': [report['Name'], row_num, 0],
                        'categories': [report['Name'], 0, 1, 0, chartend],
                        'values': [report['Name'], row_num, 1, row_num, chartend],
                    })
                chart.set_y_axis({'label_position': 'low'})
                chart.set_x_axis({'label_position': 'low'})
                worksheet.insert_chart('O2', chart, {'x_scale': 2.0, 'y_scale': 2.0})
        writer.save()

    def sendEmail(self):
        ## send email with result xls file
        if SES_SEND:
            # Email logic
            msg = MIMEMultipart()
            msg['From'] = SES_FROM
            msg['To'] = COMMASPACE.join(SES_SEND.split(","))
            msg['Date'] = formatdate(localtime=True)
            msg['Subject'] = "Cost Explorer Report"
            text = "Find your Cost Explorer report attached\n\n"
            msg.attach(MIMEText(text))
            with open("cost_explorer_report.xlsx", "rb") as fil:
                part = MIMEApplication(
                    fil.read(),
                    Name="cost_explorer_report.xlsx"
                )
            part['Content-Disposition'] = 'attachment; filename="%s"' % "cost_explorer_report.xlsx"
            msg.attach(part)
            # SES Sending
            ses = boto3.client('ses', region_name=SES_REGION)
            result = ses.send_raw_email(
                Source=msg['From'],
                Destinations=SES_SEND.split(","),
                RawMessage={'Data': msg.as_string()}
            )
            print("Successfully sent email to %s" % (SES_SEND))

def main():
    costexplorer = CostExplorer(CurrentMonth=True)
    costexplorer.setStart(args.months)
    # Overall Billing Reports
    costexplorer.addReport(Name="Total", GroupBy=[], Style='Total', IncSupport=True)

    # GroupBy Reports
    costexplorer.addReport(Name="Services", GroupBy=[{"Type": "DIMENSION", "Key": "SERVICE"}], Style='Total', IncSupport=True)

    # Generate Excel file
    costexplorer.generateExcel()

    # Send email
    costexplorer.sendEmail()
    return "Report Generated"

if __name__ == '__main__':
    main()