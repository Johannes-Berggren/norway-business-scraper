require 'Nokogiri'
require 'spreadsheet'
require 'open-uri'
require 'highline/import'

org_nr = []
contact_information = []

puts "
This program is designed to collect phone numbers from the Norwegian business catalogue Proff.no.

Currently, this program only accepts .xls files (not .xlsx). The organization numbers must be in the B column of the spreadsheet.

WARNING
This program uses your computers internet connection to collect publicly available data from Proff.
It may be illegal to use this script, especially for commercial purposes.

The author of this script definitely can't be held accountable for whatever you're using it for.

Press CTRL + C to cancel, or Enter to start scraping.
"

# Prompting user for file name until valid name is entered.
invalid_file_name = true
while invalid_file_name

  # NameError exception handling.
  begin
    filename = ask("What's the file name? (Include .xls) ")
    book = Spreadsheet.open(filename)

  rescue
    puts "Invalid file name."

  else
    sheet1 = book.worksheet(0)
    invalid_file_name = false

  end
end

# Display data to the user, and get verification

sheet1.each do |row|
  puts row[1]
end

start_scraping = ask("Press enter to start scraping. ")

sheet1.each do |row|

  # Getting orgnr from spreadsheet and searching Proff.no
  row[1] = row[1].to_i
  url = 'http://www.proff.no/bransjes%C3%B8k?q=' + row[1].to_s

  # Checking data
  data = Nokogiri::HTML(open(url))
  data.css('.tel').map do |a|
    puts a.text.gsub(/[^\d]/, '')

    # Insert result back in spreadsheet
    row.push(a.text.gsub(/[^\d]/, ''))
  end
end

book.write(filename + "-Scraped.xls")