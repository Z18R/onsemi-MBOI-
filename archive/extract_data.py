import requests
from bs4 import BeautifulSoup

# URL of the locally hosted website
url = "http://localhost:444/"

# Fetch the HTML content from the URL
response = requests.get(url)
if response.status_code == 200:
    soup = BeautifulSoup(response.text, "html.parser")
else:
    print(f"Failed to fetch the URL. Status code: {response.status_code}")
    exit()

# Extract data from the "Manufacturing Bill of Information (MBOI)" section
mboi_section = soup.find("div", class_="section")
mboi_table = mboi_section.find("table")
mboi_rows = mboi_table.find_all("tr")

mboi_data = {
    "Part Number": mboi_rows[0].find_all("td")[0].text.strip(),
    "Part Type": mboi_rows[0].find_all("td")[1].text.strip(),
    "Site | Name | Loc Code": mboi_rows[1].find_all("td")[0].text.strip(),
    "File Name": mboi_rows[1].find_all("td")[1].text.strip(),
}

# Extract data from the "Orderable Part Section - OPN" section
opn_section = mboi_section.find_next_sibling("div", class_="section")
opn_table = opn_section.find("table")
opn_rows = opn_table.find_all("tr")

opn_data = {
    "Part Number | Desc": opn_rows[0].find_all("td")[0].text.strip(),
    "Market PN": opn_rows[0].find_all("td")[1].text.strip(),
}

# Extract data from the "Assembly Part Section - ASY" section
asy_section = opn_section.find_next_sibling("div", class_="section")
asy_table = asy_section.find("table")
asy_rows = asy_table.find_all("tr")

asy_data = {
    "Part Number": asy_rows[0].find_all("td")[0].text.strip(),
    "Description": asy_rows[0].find_all("td")[1].text.strip(),
}

# Print the extracted data
print("Manufacturing Bill of Information (MBOI)")
for key, value in mboi_data.items():
    print(f"{key}: {value}")

print("\nOrderable Part Section - OPN")
for key, value in opn_data.items():
    print(f"{key}: {value}")

print("\nAssembly Part Section - ASY")
for key, value in asy_data.items():
    print(f"{key}: {value}")