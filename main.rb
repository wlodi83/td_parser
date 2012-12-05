require File.join(File.dirname(__FILE__), 'lib/exasol')
require 'spreadsheet'
require 'yaml'

config = YAML.load_file("config/config.yaml")
login = config["login"]
password = config["password"]

#Create result file
result_excel = Spreadsheet::Workbook.new
sheet1 = result_excel.create_worksheet
sheet1.name = 'Result'
sheet1.row(0).concat %w{LandingPageID Title ProgramURL TDProgramID}

#SQL queries
tradedoubler = "select lp.id, lp.title, lp.program_url from cms.affiliate_networks as an" \
  + " join cms.advertisers as adv on an.id = adv.affiliate_network_id" \
  + " join cms.programs as pr on adv.id = pr.advertiser_id" \
  + " join cms.program_regions as preg on pr.id = preg.program_id" \
  + " join cms.landing_pages as lp on preg.id = lp.program_region_id" \
  + " join cms.provisions as prov on lp.id = prov.landing_page_id" \
  + " where an.id = '8'"

row_counter = 1

#connect with exasol
connection = Exasol.new(login, password)
connection.connect
connection.do_query(tradedoubler)

#save database result in array
temp_result = []
temp_result = connection.print_result_array

temp_result.each do |column|
    if column[2].match(/(?<=p=)\d*(?=&)/)
      sheet1.row(row_counter).push column[0], column[1], column[2], column[2].match(/(?<=p=)\d*(?=&)/)[0]
    else
      sheet1.row(row_counter).push column[0], column[1], column[2], "wrong url"
    end
  row_counter += 1
end

connection.disconnect
result_excel.write 'td_result.xls'
