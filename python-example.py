import win32com.client, datetime, xlsxwriter, pandas
from datetime import datetime
from flask import *
from flask import Flask, render_template, request

app = Flask(__name__)
#from app import routes

@app.route('/', methods=['GET', 'POST'])
def form():
    return render_template('home.html')

@app.route('/hello', methods=['GET', 'POST'])
def hello():
    workbook = xlsxwriter.Workbook('Employee_wfh.xlsx')
    worksheet = workbook.add_worksheet('Employee Data')
    bold = workbook.add_format({'bold': True})

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case,
                                        # the inbox. You can change that number to reference
                                        # any other folder

    i=0
    # Start from the first cell. Rows and columns are zero indexed.
    row = 0
    col = 0

    messages = inbox.Items

    messages.Sort("ReceivedTime", True)
    message1 = messages.GetFirst()
    worksheet.write('A1', 'Date', bold)
    worksheet.write('B1', 'Name', bold)
    worksheet.write('C1', 'Email Address', bold)
    worksheet.write('D1', 'I/C Number', bold)
    worksheet.write('E1', 'Team', bold)
    worksheet.write('F1', 'Manager', bold)
    #worksheet.write('G1', 'Sick/WFH', bold)

    row+=1

    sdate = request.form['sdate']
    start_date = str(sdate)
    start_date = start_date[2:4] + start_date[5:7] + start_date[8:10]
    end_date = str(request.form['edate'])
    end_date = end_date[2:4] + end_date[5:7] + end_date[8:10]
    #row = row + 1
    t = 1

    acceptable_subjects = ["WFH", "wfh", "sick", "Sick"]
    while t == 1:
        sent_date = str(message1.SentOn)
        sent_date = sent_date[6:8] + sent_date[0:2] + sent_date[3:5]
        if sent_date >  end_date:
            continue
        elif sent_date < start_date:
            t = 0
            continue
        else:       
            subject = message1.Subject.encode('utf-8')
            if "WFH" in subject.upper() or "WORK FROM HOME" in subject.upper() or "WORKING FROM HOME" in subject.upper():
                worksheet.write(row, col, str(message1.sentOn))
                worksheet.write(row, col+1, message1.SenderName)
                #print message1.SenderName
                if message1.Class == 43:
                 if message1.SenderEmailType == "EX":
                   exUserId = str(message1.Sender.GetExchangeUser().Address).split("/")[4].split("=")[1]
                   if len(exUserId)>7:
                       if len(exUserId)>10:
                           exUserId = exUserId.split("-")[1]
                           #print len(exUserId)
                       else:
                           exUserId = exUserId[0:7]
                       #print len(exUserId)
                   worksheet.write(row, col+2, message1.Sender.GetExchangeUser().PrimarySmtpAddress)
                   worksheet.write(row, col+3, exUserId)
                   worksheet.write(row, col+4, message1.Sender.GetExchangeUser().Department)
                   #print type(message1.Sender.GetExchangeUser().GetExchangeUserManager().PrimarySmtpAddress)
                   if message1.Sender.GetExchangeUser().GetExchangeUserManager() is not None :
                       worksheet.write(row, col+5, message1.Sender.GetExchangeUser().GetExchangeUserManager().PrimarySmtpAddress)
                   else:
                       worksheet.write(row, col+5, " ")
                 else:
                   print message1.SenderEmailAddress
                   worksheet.write(row, col+2, message1.SenderEmailAddress)
                #worksheet.write(row, col+6, "WFH")
                row += 1
        message1=messages.GetNext()
        i=i+1
    workbook.close()
    df = pandas.read_excel('Employee_wfh.xlsx', sheet_name='Employee Data')
    html = df.to_html()
    #print type(html)
    str2=html[:18]+"id=testTable "+html[18:]
    #print str2
    str_csv_script = """<script>
var xport = {
  _fallbacktoCSV: true,  
  toCSV: function(tableId, filename) {
    this._filename = (typeof filename === 'undefined') ? tableId : filename;
    // Generate our CSV string from out HTML Table
    var csv = this._tableToCSV(document.getElementById(tableId));
    // Create a CSV Blob
    var blob = new Blob([csv], { type: "text/csv" });

    // Determine which approach to take for the download
    if (navigator.msSaveOrOpenBlob) {
      // Works for Internet Explorer and Microsoft Edge
      navigator.msSaveOrOpenBlob(blob, this._filename + ".csv");
    } else {      
      this._downloadAnchor(URL.createObjectURL(blob), 'csv');      
    }
  },
  _downloadAnchor: function(content, ext) {
      var anchor = document.createElement("a");
      anchor.style = "display:none !important";
      anchor.id = "downloadanchor";
      document.body.appendChild(anchor);

      // If the [download] attribute is supported, try to use it
      
      if ("download" in anchor) {
        anchor.download = this._filename + "." + ext;
      }
      anchor.href = content;
      anchor.click();
      anchor.remove();
  },
  _tableToCSV: function(table) {
    // We'll be co-opting `slice` to create arrays
    var slice = Array.prototype.slice;

    return slice
      .call(table.rows)
      .map(function(row) {
        return slice
          .call(row.cells)
          .map(function(cell) {
            return '"t"'.replace("t", cell.textContent);
          })
          .join(",");
      })
      .join("\\r\\n");
  }
};

</script>
<body style="background-color:#e7e7e7">
<h1>Team Members Working from Home</h1>
<p><button id="btnExport" onclick="javascript:xport.toCSV('testTable');"> Export to CSV</button> <em>&nbsp;&nbsp;&nbsp;Export the data to CSV</em>
  </p>"""
    open('templates/greeting.html', 'w').write(str_csv_script + str2 + "</body>")
    #print request.form['sdate']
    return render_template('greeting.html', say=request.form['sdate'], to=request.form['edate'])

if __name__ == "__main__":
    app.run()
