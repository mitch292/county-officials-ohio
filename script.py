import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

if __name__ == '__main__':
  officials = list()
  # top row of excel data
  officials.append(['county_name', 'position', 'name', 'salary', 'party', 'term_expiration'])

  ohio_homepage = requests.get("https://ohioroster.sos.state.oh.us/county_list.aspx")
  # if we get to the homepage successfully run the script
  if ohio_homepage.status_code == 200:
    soup = BeautifulSoup(ohio_homepage.content, 'html.parser')
    html = list(soup.children)[3]
    counties_table = html.find(id='MainContent_GridView1')

    counties_table_links = counties_table.find_all('a')
    counties_data_locations = list()
    # loop through all the links, each row has 2, were only concerned with internal links
    for link in counties_table_links:
      if 'county.aspx' in link['href']:
        full_url = 'https://ohioroster.sos.state.oh.us/' + link['href']
        counties_data_locations.append(full_url)
    
    # navigate to the specific counties page
    for county_url in counties_data_locations:
      county_page = requests.get(county_url)
      county_soup = BeautifulSoup(county_page.content, 'html.parser')
      county_html = list(county_soup.children)[3]
      county_body = county_html.find(id="printContent")
      county_name = county_body.find(id="MainContent_county_name").get_text()
      county_table = county_body.find(id="MainContent_GridView1")
      official_list = county_table.find_all('tr')

      # some messaging for the terminal
      print('now collecting data on', county_name) 

      # loop through the officials and add them to our list
      # we need to skip the first row and last row
      
      for row in official_list[1:-1]:
        official = list()
        official.append(county_name)
        data_points = row.find_all('td')
        
        for data_point in data_points:
          official.append(data_point.get_text())
      
        officials.append(official)

      # now we have an array of arrays with the official data

      # create an excel

      wb = Workbook()

      ws1 = wb.active
      ws1.title = 'Ohio Counties'

      print('Creating the excel document')

      for row in officials:
        ws1.append(row)

      wb.save('ohio_county_officials.xlsx')
  