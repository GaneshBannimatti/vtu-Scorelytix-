import sys
from excel_reader import read_usns_from_excel
import scraper

# -----------------------
# Get command line inputs
# -----------------------
if len(sys.argv) != 4:
    print("Usage: python scraper_excel.py <excel_file_path> <semester> <result_url>")
    sys.exit()

excel_file = sys.argv[1]
semc = sys.argv[2]
result_url = sys.argv[3]

# -----------------------
# Set URL inside scraper.py
# -----------------------
scraper.url = result_url

# -----------------------
# Read USNs from uploaded Excel
# -----------------------
usn_list = read_usns_from_excel(excel_file)

print("USNs found in Excel:", len(usn_list))

results = []

for usn in usn_list:

    print(f"\nFetching result for {usn}")

    res = scraper.fetch_result(usn)

    if res:
        results.append(res)

print("Total Results:", len(results))

# -----------------------
# Save all scraped results to Excel
# -----------------------
scraper.save_results_to_excel(results)

scraper.driver.quit()

print("🎉 EXCEL SCRAPING COMPLETED")