import requests
import json
import urllib3
import getpass
import pandas as pd
from datetime import datetime
from urllib3.exceptions import InsecureRequestWarning

# Disable SSL warning
urllib3.disable_warnings(InsecureRequestWarning)

def clean_url(apic):
    return apic.replace('https://', '').replace('http://', '')

def get_token(apic, username, password):
    apic = clean_url(apic)
    url = f"https://{apic}/api/aaaLogin.json"
    payload = {
        "aaaUser": {
            "attributes": {
                "name": username,
                "pwd": password
            }
        }
    }
    response = requests.post(url, json=payload, verify=False)
    return response.json()['imdata'][0]['aaaLogin']['attributes']['token']

def get_data(apic, token, api_path):
    apic = clean_url(apic)
    url = f"https://{apic}{api_path}"
    headers = {
        "Cookie": f"APIC-Cookie={token}"
    }
    response = requests.get(url, headers=headers, verify=False)
    return response.json()['imdata']

def main():
    # Get APIC details from user
    apic = input("Enter APIC IP/hostname: ")
    username = input("Enter username: ")
    password = getpass.getpass("Enter password: ")

    try:
        # Get authentication token
        token = get_token(apic, username, password)
        
        # Initialize data lists for Excel
        combined_data = []

        # Get Application Profiles and EPGs
        print("\n=== Gathering Application Profiles and EPGs ===")
        app_profiles = get_data(apic, token, "/api/node/class/fvAp.json")
        
        for ap in app_profiles:
            ap_name = ap['fvAp']['attributes']['name']
            tenant = ap['fvAp']['attributes']['dn'].split('/')[1][3:]
            
            # Get EPGs for this AP
            epgs = get_data(apic, token, f"/api/node/class/fvAEPg.json?query-target-filter=wcard(fvAEPg.dn,\"{ap['fvAp']['attributes']['dn']}\")")
            
            for epg in epgs:
                epg_name = epg['fvAEPg']['attributes']['name']
                epg_dn = epg['fvAEPg']['attributes']['dn']
                
                # Get BD mapping for this EPG
                bd_data = get_data(apic, token, f"/api/node/mo/{epg_dn}/rsbd.json?query-target=self")
                
                if bd_data:
                    bd_dn = bd_data[0]['fvRsBd']['attributes']['tDn']
                    bd_name = bd_dn.split('/')[-1][3:]  # Extract BD name from DN
                    
                    # Get VRF for this BD
                    vrf_data = get_data(apic, token, f"/api/node/mo/{bd_dn}/rsctx.json?query-target=self")
                    vrf_name = ''
                    if vrf_data:
                        vrf_dn = vrf_data[0]['fvRsCtx']['attributes']['tDn']
                        vrf_name = vrf_dn.split('/')[-1][3:]  # Extract VRF name from DN
                    
                    # Get Subnets from EPG level
                    epg_subnets = get_data(apic, token, f"/api/node/class/fvSubnet.json?query-target-filter=wcard(fvSubnet.dn,\"{epg_dn}\")")
                    
                    # Get Subnets from BD level
                    bd_subnets = get_data(apic, token, f"/api/node/class/fvSubnet.json?query-target-filter=wcard(fvSubnet.dn,\"{bd_dn}\")")
                    
                    # If there are subnets at either level, add them
                    if epg_subnets or bd_subnets:
                        # Add EPG level subnets
                        for subnet in epg_subnets:
                            combined_data.append({
                                'Tenant': tenant,
                                'Application Profile': ap_name,
                                'EPG': epg_name,
                                'Bridge Domain': bd_name,
                                'VRF': vrf_name,
                                'Subnet': subnet['fvSubnet']['attributes']['ip'],
                                'Subnet Level': 'EPG'
                            })
                        # Add BD level subnets
                        for subnet in bd_subnets:
                            combined_data.append({
                                'Tenant': tenant,
                                'Application Profile': ap_name,
                                'EPG': epg_name,
                                'Bridge Domain': bd_name,
                                'VRF': vrf_name,
                                'Subnet': subnet['fvSubnet']['attributes']['ip'],
                                'Subnet Level': 'BD'
                            })
                    else:
                        # If no subnets at either level, add row without subnet
                        combined_data.append({
                            'Tenant': tenant,
                            'Application Profile': ap_name,
                            'EPG': epg_name,
                            'Bridge Domain': bd_name,
                            'VRF': vrf_name,
                            'Subnet': '',
                            'Subnet Level': ''
                        })
                else:
                    # If EPG has no BD mapping, add row without BD and subnet
                    combined_data.append({
                        'Tenant': tenant,
                        'Application Profile': ap_name,
                        'EPG': epg_name,
                        'Bridge Domain': '',
                        'VRF': '',
                        'Subnet': '',
                        'Subnet Level': ''
                    })

        # Create Excel writer
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_file = f'ACI_TN-AP-EPG-BD-BDsubnet_{timestamp}.xlsx'
        
        # Convert data to DataFrame and write to Excel
        df = pd.DataFrame(combined_data)
        
        # Sort the DataFrame by Tenant, AP, EPG, BD, and Subnet
        df = df.sort_values(['Tenant', 'Application Profile', 'EPG', 'Bridge Domain', 'Subnet'])
        
        with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
            # Write the combined sheet
            df.to_excel(writer, sheet_name='ACI Inventory', index=False)

        print(f"\nData has been saved to {excel_file}")
        print(f"Total rows: {len(df)}")

    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    main()
