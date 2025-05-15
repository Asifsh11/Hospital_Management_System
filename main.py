import pandas, os, openpyxl, datetime

# download pandas,openpyxl using pip command in your cmd

ckpath = os.path.exists("Hospital_DataBase.xlsx")
if ckpath:
    wb = openpyxl.load_workbook("Hospital_DataBase.xlsx")
else:  # Create a xlsx file if it don't exist
    wb = openpyxl.Workbook()
    wb.save("Hospital DataBase.xlsx")
    sh = wb.active
    sh.title = "doctordetails"
    wb["doctordetails"].append(["DOCTOR ID", "DOCTOR NAME", "SPECILIZATION", "EXPERIENCE"])
    wb.create_sheet(title="coworkersdetails")
    wb["coworkersdetails"].append(["WORKER ID", "Name", "Qualification", "Position"])
    wb.create_sheet(title="paitentdetails")
    wb["paitentdetails"].append(["PAITENT ID", "Paitentname", "Gender", "Age", "Address", "Date&Time"])
    wb.save("Hospital DataBase.xlsx")

# Renaming the sheets with Entity
sheet1 = wb["doctordetails"]
sheet2 = wb["coworkersdetails"]
sheet3 = wb["paitentdetails"]


# Display all details of Entities
def display():
    try:
        xls = pandas.ExcelFile("Hospital DataBase.xlsx")
        data1 = pandas.read_excel(xls, "doctordetails")
        data2 = pandas.read_excel(xls, "coworkersdetails")
        data3 = pandas.read_excel(xls, "paitentdetails")
        print("\t\t1. Doctors Details\n\t\t2. Co-worker Details\n\t\t3. Patient Details")
        a = int(input("enter your choice: "))
        if a == 1:
            print(data1)
        elif a == 2:
            print(data2)
        elif a == 3:
            print(data3)
        else:
            print("Sorry, Your entered wrong choice")
    except ValueError:
        print("OPPS,You entered letters/symbols, Try again")


# Enter the details by choosing one of the entity
def enter():
    print("1. Doctors Details\n2. Co-worker Details\n3. Patient Details")
    c = int(input("Enter your Choice:"))
    if c == 1:
        did = input("\nEnter the Doctor ID\n\t").strip()
        di = input("Enter the Doctor Name\n\t").strip().upper
        dj = input("Enter the specilization\n\t").strip().capitalize()
        dk = input("Enter your experince\n\t").strip()
        sheet1.append([did, di, dj, dk])
        wb.save("Hospital DataBase.xlsx")

    elif c == 2:
        cid = input("Enter the working ID\n\t").strip()
        cn = input("Enter your Name\n\t").strip().upper
        cq = input("Enter Your Qualificationn\n\t").strip().capitalize()
        cp = input("Enter your position\n\t").strip()
        sheet2.append([cid, cn, cq, cp])
        wb.save("Hospital DataBase.xlsx")
    elif c == 3:
        ptn = input("Enter the patien token number\n\t").strip()
        pn = input("enter the paitent name\n\t").strip().upper()
        pa = input("enter the paitent age\n\t").strip()
        pg = input("enter the Gender\n\t").strip().capitalize()
        pd = input("enter the address\n\t").strip()
        dt = datetime.datetime.now()
        sheet3.append([ptn, pn, pg, pa, pd, op ,dt])
        wb.save("Hospital DataBase.xlsx")
    else:
        print("Sorry, Your entered wrong choice")


# Delete your desired person details
def delete():
    print("\t\t1. Doctors Details\n\t\t2. Co-worker Details\n\t\t3. Patient Details")
    try:
        d = int(input("enter your choice: "))
        if d == 1:
            rem = int(input("Enter the Doctor id: "))
            for cell in sheet1["A"]:
                if cell.value == rem:
                    sheet1.delete_rows(cell.row, 1)
                    print("Sucessfully deleted")
        elif d == 2:
            rem = int(input("Enter the Worker id: "))
            for cell in sheet2["A"]:
                if cell.value == rem:
                    sheet1.delete_rows(cell.row, 1)
                    print("Sucessfully deleted")
        elif d == 3:
            rem = int(input("Enter the Paitent id: "))
            for cell in sheet3["A"]:
                if cell.value == rem:
                    sheet1.delete_rows(cell.row, 1)
                    print("Sucessfully deleted")
        else:
            print("Sorry, Your entered wrong choice")
        wb.save("HMSEXCEL.xlsx")
    except ValueError:
        print("Please enter numbers only")


e = 1
print("""
        :::::::::::::::::::: Welcome to ::::::::::::::::::::
        ::::::::::::::::: Hospital DataBase ::::::::::::::::""")
print("""
    By
    Asif Shaik
    EMPIFS004984
    Government Degree Col
    lege, Rajampet
    Yogi Vemana University""")
while e != 0:
    print("""   
                1. Display the details
                2. Add a new member
                3. Delete a member
                4. Make an exit          """)
    try:
        b = int(input("Enter your Choice:"))
        if b == 1:
            display()
        elif b == 2:
            enter()
        elif b == 3:
            delete()
        elif b == 4:
            e = 0
        else:
            print("You entered wrong choice, Please Try Again")
    except ValueError:
        print("You entered wrong choice, Please Try Again")
