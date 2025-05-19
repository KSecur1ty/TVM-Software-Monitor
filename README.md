# TVM-Software-Monitor
A python command line tool used in conjunction with Defender Hunting Queries to monitor when software on managed endpoints changes state, either installed or uninstalled software and records them in two separate excel files for ease.  
![TVM Logo](https://github.com/user-attachments/assets/8584340c-2367-45cf-8bd9-d0462775f4bb)


Defender Advanced Hunting with KQL (Kusto Query Language) is incredibly powerful when used in conjuction with other tools, especially regarding data about Endpoints you manage within your Microsoft Tenant. TVM stands for Threat and Vulnerabilities Management.
By using the below query, we can collect all the data from defender's DeviceTvmSoftwareInventory to see a per device view of each endpoints installed programs they have on it. It is important to note that the query does NOT show default office products when looking at Windows devices. It can however detect applications regardless of the endpoints operating system.

****DeviceTvmSoftwareInventory
| extend SoftwareInfo = strcat(SoftwareName, " (", SoftwareVersion, ")")
| summarize InstalledSoftware = make_list(SoftwareInfo) by DeviceName
| order by DeviceName asc****


Once this query has finished running, navigate to export and press download as CSV, before saving it to the same location as this repository.
![image](https://github.com/user-attachments/assets/9da72f00-4f78-4925-9476-ec0062f5ee62)

Make sure to have Python 3.12 or later installed, before running '''pip install openpyxl'''
