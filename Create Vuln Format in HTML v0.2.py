import pandas as pd
vuln = pd.read_excel("Unique_Vul_CVE.xlsx")
vuln.head()
css = """
<style type="text/css">
.tg  {border-collapse:collapse;border-spacing:0;}
.tg td{font-family:Calibri, sans-serif;font-size:14px;padding:10px 5px;border-style:solid;border-width:1px;overflow:hidden;word-break:normal;}
.tg th{font-family:Calibri, sans-serif;font-size:14px;font-weight:normal;padding:10px 5px;border-style:solid;border-width:1px;overflow:hidden;word-break:normal;}
.tg .tg-lqbj{font-weight:bold;background-color:#d8d6d6;vertical-align:top}
.tg .tg-xdvw{font-weight:bold;font-size:20px;background-color:#d8d6d6;vertical-align:top}
.tg .tg-5yj4{background-color:#d8d6d6;vertical-align:top}
.tg .tg-n7v2{font-weight:bold;font-size:20px;background-color:#d8d6d6;vertical-align:top}
.tg .tg-s676{font-weight:bold;font-size:20px;background-color:#e6e4e4;color:#fe0000;vertical-align:top}
.tg .tg-yw4l{vertical-align:top}
</style>"""
table = """
<table class="tg" width = 800>
  <tr>
    <th class="tg-n7v2" width = 20>{No}</th>
    <th class="tg-xdvw">{Name}</th>
    <th class="tg-5yj4" colspan="2" width = 40>Level of Risk</th>
    <th class="tg-s676" width = 60>{Risk}</th>
  </tr>
  <tr>
    <td class="tg-yw4l" colspan="5"></td>
  </tr>
  <tr>
    <td class="tg-lqbj" colspan="5">Observations</td>
  </tr>
  <tr>
    <td class="tg-yw4l" colspan="5" rowspan="2">{Description}</td>
  </tr>
  <tr>
  </tr>
  <tr>
    <td class="tg-lqbj" colspan="5">Recommended Risk Mitigation</td>
  </tr>
  <tr>
    <td class="tg-yw4l" colspan="5">{Solution}</td>
  </tr>
  <tr>
    <td class="tg-lqbj" colspan="5">CVE</td>
  </tr>
  <tr>
    <td class="tg-yw4l" colspan="5">{CVE}</td>
  </tr>
  <tr>
    <td class="tg-lqbj" colspan="5">Locations</td>
  </tr>
  <tr>
    <td class="tg-yw4l" colspan="2" width = 740>{IP}</td>
    <td class="tg-yw4l" colspan="3" width = 60>{Location}</td>
  </tr>
</table>
<br><br>"""
f = open('Test write HTML.html', 'w')
f.write(css)

for number, row in vuln.iterrows():
    ftable = table.format(No=row.No, Name=row.Name, Description=row.Description, Solution=row.Solution, CVE=row.CVE, IP=row.IP, Location=row.Location, Risk=row.Risk)
    f.write(ftable)
f.close()
