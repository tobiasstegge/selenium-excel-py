import requests
import json
import pandas as pd

url = 'http://ias.vdma.org/members?p_p_lifecycle=2&p_p_resource_id=getPage&p_p_id=vdma2publicusers_WAR_vdma2publicusers&s=&page=2'
df = pd.DataFrame(columns=['Unternehmen', 'Mitarbeiter', 'Homepage', 'Info', 'Kontaktperson'])


for page in range(1, 345):
    session = requests.Session()
    data = {
      '_org_vdma_publicusers_portlet_PublicUsersPortlet_INSTANCE_H0VO3QljCiRM_query': '',
      '_org_vdma_publicusers_portlet_PublicUsersPortlet_INSTANCE_H0VO3QljCiRM_s': '',
      '_org_vdma_publicusers_portlet_PublicUsersPortlet_INSTANCE_H0VO3QljCiRM_page': str(page)
    }
    request = requests.post(url, data=data)
    content = request.content

    print(content)

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('fetch_vdma.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
df.to_excel(writer, sheet_name='Sheet1')

# Close the Pandas Excel writer and output the Excel file.
writer.save()

