from openpyxl import load_workbook

def perf():
    workbook = load_workbook("wsheet.xlsx")
    currentwork = workbook[workbook.sheetnames[0]]

    f = open("test.html" , "r")
    lines = f.readlines()
    nlines = len(lines) #count the number of lines
    f = open("test.html" , "w")
    for i in range(0 , nlines-4):
        f.write(lines[i])              #cursor reaches to the line where we need to edit
    f.write("\n")
    a2=currentwork["A2"]
    f.write("<a href=\""+a2.value+".html\">\n")
    f.write("<div id=\"b3\">")
    f.write("<img src=\""+a2.value+".png\" width = \"180\" hesight = \"240\">\n")
    f.write("<h5>"+a2.value+"</h5>\n")
    b2 = currentwork["B2"]
    f.write("<span>"+b2.value+"</span>\n")
    c2 = currentwork["C2"]
    f.write("<div> &#x20b9; "+str(c2.value)+"</div></div></a>\n")
    f.write(lines[-4])
    f.write(lines[-3])
    f.write(lines[-2])
    f.write(lines[-1])
    currentwork.delete_rows(2, 1)

    
    f.close()
    workbook.save("wsheet.xlsx")

while ((load_workbook("wsheet.xlsx")[load_workbook("wsheet.xlsx").sheetnames[0]]["A2"]).value != None ):
    perf()


