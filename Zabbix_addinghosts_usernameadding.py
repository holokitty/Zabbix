import pandas as pd
from pyzabbix import ZabbixAPI

# File location
excel_file = input("Input the file way: ")

# Server Zabbix address
zabbix_server = input("Input Zabbix server address: ")

# Input login and password
zabbix_user = input("Input login: ")
zabbix_password = input("Input password: ")

# Connect to Zabbix API
zapi = ZabbixAPI(zabbix_server)
zapi.login(zabbix_user, zabbix_password)

# Receive information from exel file
def add_hosts_from_excel(excel_file):
    data = pd.read_excel(excel_file, dtype={'groupid': str})

    for index, row in data.iterrows():
        host_name = row['host']
        ip_address = row['ip']
        host_group = str(row['groupid'])
        # Convert to str if it is number
        group_list = [{'groupid': group_id} for group_id in host_group.split(",")]

        # Receive info groupid2
        host_group2 = str(row['groupid2'])
        group_list2 = [{'groupid': group_id} for group_id in host_group2.split(",")]

        # Adding second group to group list
        group_list.extend(group_list2)

        # Receive templates id from exel
        templatesid = row['template_id']

        # Receive proxy_id frome exel
        proxy_id = row['proxy_id']

        #Adding tags from exel
        tags = [{'tag': f'{row["Tag1"]}', 'value': f'{row["Tag1Value"]}'},
                {'tag': f'{row["Tag2"]}', 'value': f'{row["Tag2Value"]}'},
                {'tag': f'{row["Tag3"]}', 'value': f'{row["Tag3Value"]}'}]

        template_list = [{'templateid': templatesid}] if pd.notnull(templatesid) else []
        proxy_hostid = proxy_id if pd.notnull(proxy_id) else None

        host = zapi.host.create(
            host=host_name,
            interfaces=[{
                'type': 1,
                'main': 1,
                'useip': 1,
                'ip': ip_address,
                'dns': '',
                'port': '10050'
            }],
            groups=group_list,
            tags=tags,
            templates=template_list,
            proxy_hostid=proxy_hostid
        )

        print(f"Host '{host_name}' added with host ID: {host['hostids'][0]}")

# The function add_hosts_from_excel()
excel_path = '/Users/mariiagit/venv/python_codes/Zabbix.xlsx'
add_hosts_from_excel(excel_path)

# Close connect
zapi.logout()
