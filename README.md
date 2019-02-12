# ServeExcel
Serves Excel files uploaded via web commands.

API Build report template Columns

<Parameter> reportname: name of report to create or update

<Parameter> header: comma delimenated column names.

<Parameter> createnew: true/false to create or update the report

the below example will create a new Template Object with the format set to values passed to the header parameter

http://127.0.0.1:8183/?reportname=DynamicReport&header=field1,field2,field3,field4&createnew=true


the next example matches the existing report by name and updates the header to the new values passed

http://127.0.0.1:8183/?reportname=DynamicReport&header=column1,column2,column3,column4&createnew=false

API Add report data

<Parameter> /?reportname=nameOfReport: name of report to add content to

<Parameter> &content=: comma delimenated column values

example

http://127.0.0.1:8183/?reportname=DynamicReport&content=1,2,3,4,a,b,c,d

API GET resulting report by name

<Parameter> /?getreport=nameOfReport

example

http://127.0.0.1:8183/?getreport=DynamicReeport

API Query available reports

<Parameter> "/?reports" : basic request to list available reports

example

http://127.0.0.1:8183/?reports


Next steps

    1. Decentralize the REPORTS dataset so that multi reports in 1 excel file can be achieved.(started)

    2 serialize to memory instead of files for templates.(unstarted)

    3. excel data generator in done by data type formet with constrictions on  content of data generated.(unstarted)

    4. Convert completely to .net core or a .net standard libary.(unstarted)

Pull requests are welcome. :)

