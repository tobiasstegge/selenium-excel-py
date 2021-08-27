import requests
import json
import pandas as pd

url = 'https://www.vdma.org/mitglieder?p_p_id=org_vdma_publicusers_portlet_PublicUsersPortlet_INSTANCE_H0VO3QljCiRM&p_p_lifecycle=2&p_p_state=normal&p_p_mode=view&p_p_resource_id=getPage&p_p_cacheability=cacheLevelPage'
df = pd.DataFrame(columns=['Unternehmen', 'Mitarbeiter', 'Homepage', 'Info', 'Kontaktperson'])


for page in range(1, 345):
    # Get data
    data = {
      '_org_vdma_publicusers_portlet_PublicUsersPortlet_INSTANCE_H0VO3QljCiRM_query': '',
      '_org_vdma_publicusers_portlet_PublicUsersPortlet_INSTANCE_H0VO3QljCiRM_s': '',
      '_org_vdma_publicusers_portlet_PublicUsersPortlet_INSTANCE_H0VO3QljCiRM_page': str(page)
    }
    session = requests.Session()
    request = requests.post(url, data=data)
    response = request.json()
    companies = json.loads(response['publicUserList'])['content']

    # Filter all not from germany
    websites = []
    for company in companies:
        if company['country'] == 'Deutschland' and company['webAddr'] not in websites and company['companyName'] != '':
            websites.append(company['webAddr'])
            df = df.append({
                'Unternehmen': company.get('companyName'),
                'Mitarbeiter': '',
                'Homepage': company.get('webAddr'),
                'Info': f"Stadt: {company.get('city', '')}",
                'Kontaktperson': f"Tel: {company.get('phoneNum', '')} \n Email: {company.get('email', '')}"
            }, ignore_index=True)
        else:
            print(f"Sorted out {company.get('companyName')}")

    print(page)

# Write to excel
writer = pd.ExcelWriter('fetch_vdma.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='Sheet1')
writer.save()

