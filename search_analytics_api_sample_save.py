"""Example for using the Google Search Analytics API (part of Search Console API).
A basic python command-line example that uses the searchAnalytics.query method
of the Google Search Console API. This example demonstrates how to query Google
search results data for your property. Learn more at
https://developers.google.com/webmaster-tools/
To use:
1) Install the Google Python client library, as shown at https://developers.google.com/webmaster-tools/v3/libraries.
2) Sign up for a new project in the Google APIs console at https://code.google.com/apis/console.
3) Register the project to use OAuth2.0 for installed applications.
4) Copy your client ID, client secret, and redirect URL into the client_secrets.json file included in this package.
5) Run the app in the command-line as shown below.
Sample usage:
  $ python search_analytics_api_sample_save.py https://www.unicefusa.org/
"""
import re
import argparse
import datetime
import xlwt
import sys
from googleapiclient import sample_tools

# Declare command-line flags.
argparser = argparse.ArgumentParser(add_help=False)
argparser.add_argument('property_uri', type=str,
                       help=('Site or app URI to query data for (including '
                             'trailing slash).'))


def main(argv):
  service, flags = sample_tools.init(
      argv, 'webmasters', 'v3', __doc__, __file__, parents=[argparser],
      scope='https://www.googleapis.com/auth/webmasters.readonly')

  report = xlwt.Workbook()
  sheet1 = report.add_sheet('Sheet 1')
  sheet1.write(0, 0, str(datetime.date.today()))
  sheet1.write(0, 1, 'Impressions')
  sheet1.write(0, 2, 'Clicks')
  sheet1.write(0, 3, 'CTR')
  xlcnt=1
  # Set date for this week.
  s_date = datetime.date.today()-datetime.timedelta(days=9)
  e_date = datetime.date.today()-datetime.timedelta(days=3)
  
  #Build, execute, print request for this week
  request = build_request(s_date, e_date)
  response = execute_request(service, flags.property_uri, request)
  print_table(response, 'This year requests '+str(s_date)+' '+str(e_date), xlcnt, sheet1)
  xlcnt += 1
  # Set date for previous week.
  ws_date = s_date - datetime.timedelta(days = 7)
  we_date = e_date - datetime.timedelta(days = 7)
  
  #Build, execute, print request for previous week
  request = build_request(ws_date, we_date)
  response = execute_request(service, flags.property_uri, request)
  print_table(response, 'This year previous week requests '+str(ws_date)+' '+str(we_date), xlcnt, sheet1)
  xlcnt += 1
  
  # Set date for previous year.
  ys_date = s_date - datetime.timedelta(days = 364)
  ye_date = e_date - datetime.timedelta(days = 364)

  #Build, execute, print request for previous year
  request = build_request(ys_date, ye_date)
  response = execute_request(service, flags.property_uri, request)
  print_table(response, 'Previous year requests '+str(ys_date)+' '+str(ye_date), xlcnt, sheet1)
  report.save('SEO_UNICEF_report.xls')
  
def build_request(start_date, end_date):
  request = {
      'startDate': str(start_date),
      'endDate': str(end_date),
      'dimensions': ['query'],
      'rowLimit': 25000
  }
  return request

def execute_request(service, property_uri, request):
  """Executes a searchAnalytics.query request.
  Args:
    service: The webmasters service to use when executing the query.
    property_uri: The site or app URI to request data for.
    request: The request to be executed.
  Returns:
    An array of response rows.
  """
  return service.searchanalytics().query(
      siteUrl=property_uri, body=request).execute()


def print_table(response, title, xlcnt, sheet1):
  
  totclicks = 0
  totimpr = 0
  totctr = 0.0
  """Prints out a response table.
  Each row contains key(s), clicks, impressions, CTR, and average position.
  Args:
    response: The server response to be printed as a table.
    title: The title of the table.
  """
  sheet1.write(xlcnt, 0, title + ':')
  print (title + ':')

  if 'rows' not in response:
    print ('Empty response')
    sheet1.write(xlcnt, 1, 'Empty response')
    return
  row_cnt = 0
  rows = response['rows']
  row_format = '{:<20}' + '{:>20}' * 3
  print (row_format.format('Keys', 'Clicks', 'Impressions', 'CTR'))
  
  for row in rows:
    row_cnt +=1 
    regcheck = ''
    skey = str(row['keys'])
    regcheck = re.search(r'unicef|un...f|un..f|un....f', skey)
    if regcheck is not None:
      totclicks += int(row['clicks'])
      totimpr += int(row['impressions'])
  totctr = totclicks/totimpr
  print (row_format.format('Unicef regex filter', totclicks, totimpr, totctr))
  sheet1.write(xlcnt, 1, totimpr)
  sheet1.write(xlcnt, 2, totclicks)
  sheet1.write(xlcnt, 3, totctr)
  print(row_cnt)

if __name__ == '__main__':
  main(sys.argv)