import pandas as pd
all_vuln = pd.read_excel("All_vuln.xlsx")
all_vuln.head()
uvuln = pd.read_excel("Unique_Vuln.xlsx")
uvuln.head()
names_ip_di = {name:set()
              for name in all_vuln["Name"]}
for name, host, i in zip(all_vuln['Name'], all_vuln['Host'], range(len(all_vuln))):
        names_ip_di[name].add(host)
names_ip_di = {name:sorted(ips) for name, ips in names_ip_di.items()}
uvuln["IP"] = [names_ip_di[name] for name in uvuln["Name"]]
uvuln.head()
uvuln.sort_index().to_csv('Vuln_sorted_hosts.csv')
