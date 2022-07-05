import mysql.connector as mysql
from tabulate import tabulate
from win32com.client import Dispatch
from playsound import playsound


#variables
options = ['View records', 'View categories', 'View branches', 'Add record', 'Edit record', 'Delete record']
options2 = ['Name', 'Category', 'Branch', 'Bill to be paid', 'Amount already paid', 'Notes']

#text to speech
speak = Dispatch("SAPI.SpVoice").Speak

#establishes connection with mysql
def startup():
    while True:
        speak("Please enter your details")
        user = input('Please enter your mysql username: ')
        passwd = input('Please enter your mysql password: ')
        try:
            myconn = mysql.connect(host='localhost', user=user, passwd=passwd)
            if myconn.is_connected():
                print('*****Connection established*****')
                cursor = myconn.cursor()
                return cursor, myconn
        except:
            print('Invalid credentials!')
        
#creates required database        
def setup(cursor, myconn):
    
    #create database
    try:
        cursor.execute('create database accounts;')
    except:
        pass
    
    #use database
    try:
        cursor.execute('use accounts;')
        print('*****Database selected*****')
    except:
        pass
    
    #create tables
    try:
        cursor.execute('create table records(s_no INT(3) AUTO_INCREMENT PRIMARY KEY, client VARCHAR(40) NOT NULL, category_id INT(3) NOT NULL, branch_id INT(3) NOT NULL, bill INT(8) NOT NULL, paid INT(8), notes VARCHAR(80));')
        cursor.execute('create table category(category_id INT NOT NULL PRIMARY KEY, category_name VARCHAR(30) NOT NULL);')
        cursor.execute('create table branch(branch_id INT NOT NULL PRIMARY KEY, branch_name VARCHAR(30) NOT NULL);')
        cursor.execute('insert into branch values(125, "Dubai");')
        cursor.execute('insert into branch values(126, "Sharjah");')
        cursor.execute('insert into branch values(127, "Abu Dhabi");')
        cursor.execute('insert into category values(371, "Energy");')
        cursor.execute('insert into category values(372, "Gas")')
        cursor.execute('insert into category values(373, "Fire");')
        cursor.execute('insert into category values(374, "Construction");')
        print('*****Tables successfully setup*****')
        myconn.commit()
    except:
        pass
    
#menu driven program
def run(cursor, myconn):
    while True:
        try:
            speak('Select one of the listed options')
            print('********************************************************************')
            for i in range(len(options)):
                print(f'{i+1}: {options[i]}')
            print('0: Exit system')
            userinput = int(input("Please select one of the above options: "))
            print('********************************************************************')
            if userinput == 1:
                viewrecords(cursor, myconn)
            elif userinput == 2:
                viewcategories(cursor, myconn)
            elif userinput == 3:
                viewbranches(cursor, myconn)
            elif userinput == 4:
                addr(cursor, myconn)
            elif userinput == 5:
                editr(cursor, myconn)
            elif userinput == 6:
                remover(cursor, myconn)
            elif userinput == 0:
                speak('Exiting system')
                print('All changes saved!')
                playsound('./beep.mp3')
                break
            else:
                speak('Please enter a valid option')
        except:
            speak('Please enter a valid option')

#function to add records
def addr(cursor, myconn):
    print("Enter the following details as prompted below:")
    name = input("Client's name - ")
    cat = int(input("Category ID - "))
    branch = int(input("Branch ID - "))
    bill = int(input("Amount to be billed - "))
    paid = int(input("Amount already paid, if any - "))
    notes = input("Additional notes - ")
    cursor.execute("insert into records(client, category_id, branch_id, bill, paid, notes) values('{0}', {1}, {2}, {3}, {4}, '{5}');".format(name, cat, branch, bill, paid, notes))
    myconn.commit()

# function to view categories
def viewcategories(cursor, myconn):
    cursor.execute('select * from category;')
    data = cursor.fetchall()
    categories = tabulate(data, headers=['ID', 'Category'], tablefmt='pretty')
    print()
    print(categories)
    print()
    print('********************************************************************')
    input("Press Enter to continue...")

# function to view records
def viewrecords(cursor, myconn):
    cursor.execute('select * from records;')
    data = cursor.fetchall()
    records = tabulate(data, headers=['ID', 'Client', 'Category_ID', 'Branch_ID', 'Bill', 'Paid', 'Notes'], tablefmt='pretty')
    print()
    print(records)
    print()
    print('********************************************************************')
    input("Press Enter to continue...")

    
# function to view branches
def viewbranches(cursor, myconn):
    cursor.execute('select * from branch;')
    data = cursor.fetchall()
    branches = tabulate(data, headers=['ID', 'Branch'], tablefmt='pretty')
    print()
    print(branches)
    print()
    print('********************************************************************')
    input("Press Enter to continue...")

#function edit values
def editr(cursor, myconn):
    cursor.execute('select * from records;')
    data = cursor.fetchall()
    records = tabulate(data, headers=['ID', 'Client', 'Category_ID', 'Branch_ID', 'Bill', 'Paid', 'Notes'], tablefmt='pretty')
    print()
    print(records)
    print()
    print('********************************************************************')
    serial = input('Select the index of the entry to be edited: ')
    for i in range(len(options2)):
        print(f'{i+1}: {options2[i]}')
    col = int(input('Select the attribute to be edited: '))
    col -= 1
    if col in range(len(options2)):
        newval = input(f'Please enter the new value for the selected column: ')
        if col == 0:
            cursor.execute(f'update records set client = "{newval}" where s_no = {int(serial)}')
            myconn.commit()
            print('Update completed!')
            speak('Update completed!')
        elif col == 1:
            cursor.execute(f'update records set category_id = {int(newval)} where s_no = {int(serial)}')
            myconn.commit()
            print('Update completed!')
            speak('Update completed!')
        elif col == 2:
            cursor.execute(f'update records set branch_id = {int(newval)} where s_no = {int(serial)}')
            myconn.commit()
            print('Update completed!')
            speak('Update completed!')
        elif col == 3:
            cursor.execute(f'update records set bill = {int(newval)} where s_no = {int(serial)}')
            myconn.commit()
            print('Update completed!')
            speak('Update completed!')
        elif col == 4:
            cursor.execute(f'update records set paid = {int(newval)} where s_no = {int(serial)}')
            myconn.commit()
            print('Update completed!')
            speak('Update completed!')
        elif col ==5:
            cursor.execute(f'update records set notes = "{newval}" where s_no = {int(serial)}')
            myconn.commit()
            print('Update completed!')
            speak('Update completed!')
    else:
        pass

def remover(cursor, myconn):
    cursor.execute('select * from records;')
    data = cursor.fetchall()
    records = tabulate(data, headers=['ID', 'Client', 'Category_ID', 'Branch_ID', 'Bill', 'Paid', 'Notes'], tablefmt='pretty')
    print()
    print(records)
    print()
    print('********************************************************************')
    delindex = input("Kindly select the serial number of the record you'd like to delete: ")
    cursor.execute(f'delete from records where s_no = {delindex};')
    myconn.commit()
    print('Successfully deleted!')

#main, runs all functions required for the file
def mainfunc():
    playsound('./beep.mp3')
    cursor, myconn = startup()
    setup(cursor, myconn)
    speak('Welcome to the client management system.')
    print('Welcome!')
    run(cursor, myconn)

# calls the main function
mainfunc()