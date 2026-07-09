<%
' Copy this file to datastore.asp and fill in local values.
Response.Buffer = True

strconnect = "Provider=SQLOLEDB; Data Source=YOUR_SQL_SERVER; Initial Catalog=YOUR_DATABASE; User Id=YOUR_USER; Password=YOUR_PASSWORD"
orderexcelpath = "D:\path\to\shared"
INvoiceType = "INVS"
uplopath = "/your-app/shared/"
documentpath = "shared/"
documentpath_claims = "shared/claims/"

CompanyHeaderName = "YOUR_COMPANY_NAME"
CompanyHeaderReg = "YOUR_COMPANY_REGISTRATION"
CompanyHeaderAddress = "YOUR_COMPANY_ADDRESS"
CompanyHeaderTel = ""
CompanyHeaderWeb = ""
CompanyName = "YOUR_COMPANY_NAME"
prefixname = "Live"

GSTCode = "SR"
GSTRate = 0
GSTRateBack = 0

' Configure SMTP locally in datastore.asp if email sending is enabled.
%>
