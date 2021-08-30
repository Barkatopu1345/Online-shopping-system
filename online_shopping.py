import openpyxl 

productID = 1000  



def giveID():
    path = "shopingFile.xlsx"
    wb_obj = openpyxl.load_workbook(path)
    sheet = wb_obj['items']
    lst = []
    
    __row__ = sheet.max_row
    
    for i in range(2,__row__):
        x = sheet.cell(i,1).value
        lst.append(x)
    return lst



def createAccount():
    path = "shopingFile.xlsx"
    wb_obj = openpyxl.load_workbook(path)
    sheet = wb_obj['userAccount']
    row = sheet.max_row

    lst = []
    for i in range(2,row+1):
        lst.append(sheet.cell(i,2).value)

    name = input("Enter name: ")
    userID = input("UserId(name+digits. Ex: barkat1345): ")
    password = input("Choose a password: ")
    address = input("Enter your address: ")
    phone = int(input("Enter your phone number: "))
    if userID in lst:
        print("User id can not be accepted! please try with another user id!")
        createAccount()

    else:



        lst2 = []
        sheet['A1'] = "Name"
        sheet['B1'] = "User ID"
        sheet['C1'] = "Password"
        sheet['D1'] = "Address"
        sheet['E1'] = "Phone"

        sheet.cell(row+1,1,value = name)
        sheet.cell(row+1,2,value = userID)
        sheet.cell(row+1,3,value = password)
        sheet.cell(row+1,4,value = address)
        sheet.cell(row+1,5,value = phone)

        lst2.append(name)
        lst2.append(userID)
        lst2.append(password)
        lst2.append(address)
        lst2.append(phone)
        print("Congratulations! you have created an account. ")
        wb_obj.save("shopingFile.xlsx")
        return lst2
    wb_obj.close()
     

def logIn():
    wb_obj = openpyxl.load_workbook("shopingFile.xlsx")
    sheet = wb_obj['soldProduct']
    sheet2 = wb_obj['userAccount']

    row = sheet.max_row
    column = sheet.max_column

    userId = "0"
    password = "0"
    temp = []
    userId = input("Enter your user id: ")
    password = input("Enter your password: ")

    lst = []

    for i in range(2,row+2):
        lst.append(sheet2.cell(i,2).value)
        lst.append(sheet2.cell(i,3).value)
    if userId not in lst:
        print("Please enter valid user id or password! ")
        logIn()
    else:
        index = (lst.index(userId))+1
        
        for i in range(1,6):
            temp.append(sheet2.cell(index,i).value)
    



    return temp

    

def buyProducts():
    path = "shopingFile.xlsx"
    wb_obj = openpyxl.load_workbook(path)
    sheet = wb_obj['items']

    row = sheet.max_row

    printItems()

    productID = int(input("Enter product ID: "))
    quantity = int(input("Enter quantity: "))


    productName = "0"
    price = 0
    total = 0
    lst = []
    for i in range(1,row+1):
        lst.append(sheet.cell(i,1).value)
    if productID in lst:
        index = (lst.index(productID))+1
        if(sheet.cell(index,4).value<quantity):
            print("Does not stock this much products!")
            buyProducts()
        else:
            productName = sheet.cell(index,2).value
            price = sheet.cell(index,3).value
            total =sheet.cell(index,3).value * quantity
            stockQuantity = sheet.cell(index,4).value 

            sheet.cell(index,4,value = stockQuantity - quantity) 
    else:
        print("Invalid product id. Please reenter product id again!")
        buyProducts()

    wb_obj.save("shopingFile.xlsx")

    decission = input("Complete your parchase? y/n: ")
    if decission == 'y' or decission == 'Y':
        confirmBuy(productID,productName,price,quantity,total)


    wb_obj.close()



def confirmBuy(productID,productName,price,quantity,total):
    
    path = "shopingFile.xlsx"
    wb_obj = openpyxl.load_workbook(path)
    sheet = wb_obj['soldProduct']

    row = sheet.max_row
    column = sheet.max_column
    
    profile = logIn()
    # print(profi)
    name = profile[0]
    address = profile[3]
    phone = profile[4]
    method = input("Enter payment method: cash on delivary(COD)/bkash/nogod: ")    

    sheet['A1'] = "Product ID"
    sheet['B1'] = "Product Name"
    sheet['C1'] = "Quantity"
    sheet['D1'] = "Unit price"
    sheet['E1'] = "Total"
    sheet['F1'] = "Name"
    sheet['G1'] = "address"
    sheet['H1'] = "phone"
    sheet['I1'] = "Payment Method"

    print("Product ID "+" Product Name "+" Quantity "+" Unit price "+ " Total "+" Name "+" address "+" phone "+" Payment Method ")
    print(str(productID)+"        "+productName+"       "+str(quantity)+"  "+str(price)+"  "+str(total)+"  "+name+"  "+address+"  "+str(phone)+" "+method)
   
    decission = input("Confirm your order? y/n: ")
    if decission == 'y' or decission == 'Y':
        
        sheet.cell(row+1,1,value = productID)
        sheet.cell(row+1,2,value = productName)
        sheet.cell(row+1,3,value = quantity)
        sheet.cell(row+1,4,value = price)
        sheet.cell(row+1,5,value = total)
        sheet.cell(row+1,6,value = name)
        sheet.cell(row+1,7,value = address)
        sheet.cell(row+1,8,value = str(phone))
        sheet.cell(row+1,9,value = method)

    wb_obj.save("shopingFile.xlsx")
    wb_obj.close()
    print("Thank you for staying with us! ")



def writeItems():
    path = "shopingFile.xlsx"
    wb_obj = openpyxl.load_workbook(path)
    sheet = wb_obj['items']
    printItems()
    sheet['A1'] = "Product ID"
    sheet['B1'] = "Product Name"
    sheet['C1'] = "Unit price"
    sheet['D1'] = "Quantity"

    __row__ = sheet.max_row
    column = sheet.max_column
    
    if __row__<2:
        id = productID
    else: id =  sheet.cell(__row__,1).value
    
    i = __row__+1
    loopChecker = True
    while(loopChecker == True):
        name = input("Product Name: ")
        
        name = name.capitalize()


        if (checkExistingProduct(name) == True):
            x = input("Product already exist in the list!")
            loopChecker = False
            break

            
                
        elif checkExistingProduct(name) == False:        
            price = float(input("Enter product unit price(BDT): "))
            quantity = float(input("Quantity: "))
            id += 1
            sheet.cell(i,1,value = id)
            lst = giveID()


            sheet.cell(i,2,value = name)
            
            
            sheet.cell(i,3,value = price)
            
            sheet.cell(i,4,value = quantity)
            check = input("Want to add more items? y/n: ")
            i += 1
            print(end = '\n')
            if check == 'n' or check == 'N':
                loopChecker = False

    
    wb_obj.save('shopingFile.xlsx') 
    printItems()
    wb_obj.close()



def checkExistingProduct(name):
    path = "shopingFile.xlsx"
    wb_obj = openpyxl.load_workbook(path)
    sheet = wb_obj['items']
    __row__ = sheet.max_row
    lst = []
    for i in range(1,__row__+1):
        lst.append(sheet.cell(i,2).value)
    if name in lst:
        return True
    else: return False
    wb_obj.close()


def printItems():
    path = "shopingFile.xlsx"
    wb_obj = openpyxl.load_workbook(path)
    sheet = wb_obj['items']
    __row__ = sheet.max_row
    column = sheet.max_column
    print(sheet.cell(1,1).value+"       "+sheet.cell(1,2).value+"       "+sheet.cell(1,3).value+"       "+sheet.cell(1,4).value)
    for i in range(2,__row__+1):
        for j in range(1,column+1):
            print(sheet.cell(i,j).value, end = "            ")
        print()

def showItems():
    path = "shopingFile.xlsx"
    wb_obj = openpyxl.load_workbook(path)
    sheet = wb_obj['items']
    __row__ = sheet.max_row
    column = sheet.max_column
    print(sheet.cell(1,1).value+"       "+sheet.cell(1,2).value+"       "+sheet.cell(1,3).value+"       "+sheet.cell(1,4).value)
    for i in range(2,__row__+1):
        for j in range(1,column+1):
            print(sheet.cell(i,j).value, end = "            ")
        print()



def updateItems():
    path = "shopingFile.xlsx"
    wb_obj = openpyxl.load_workbook(path)
    sheet = wb_obj['items']
    __row__ = sheet.max_row
    column = sheet.max_column
    
    lst = []
    for i in range(2,__row__+1):
        x = sheet.cell(i,1).value
        lst.append(x)
    printItems()
    check = True
    while(check == True):
        id = int(input("Enter your product ID:" ))
        id = (lst.index(id))+2
        print("check 03")
        
        name = input("Product Name: ")
        
        sheet.cell(id,2,value = name)
        
        price = float(input("Enter product unit price(BDT): "))
        
        sheet.cell(id,3,value = price)
        quantity = float(input("Quantity: "))
        
        sheet.cell(id,4,value = quantity)
        wb_obj.save('shopingFile.xlsx')

        check = input("Want to update more? y/n: ")
        if check == 'y' or check == 'Y':
            updateItems() 
        else: check == False
    
    
    printItems() 
    wb_obj.close()


def deleteItems():
    path = "shopingFile.xlsx"
    wb_obj = openpyxl.load_workbook(path)
    sheet = wb_obj['items']

    __row__ = sheet.max_row
    printItems()
    id = int(input("Enter your product ID: "))
    index = 0
    for i in range(1,__row__+1):
        if sheet.cell(i,1).value == id:
            index = i
    print(index)
    sheet.delete_rows(index,1)
    wb_obj.save('shopingFile.xlsx')
    printItems()
    wb_obj.close()




def Menu():
    decission = 0
    account = input("Do you have any account? y/n: ")
    if account == 'y' or account == 'Y':
        
        existAccount = input("Do you want to login? y/n: ")
        
        if existAccount == 'y' or existAccount == 'Y':
            profile = []
            profile = logIn()
            if (profile[1] == 'barkat1345' and profile[2] == '1234'):
                print("Add new item in the list -> 1")
                print("Update an item in the list -> 2")
                print("Delete an item from the list -> 3")
                decission = int(input("Chose your option: "))
            else:
                print("Buy a product -> 4")
                decission = int(input("Chose your option: "))
        else:
            print("Thank you for staying with us! See you later")
            return
    else:
        print("Creating new account!") 
        createAccount() 
        Menu()   
                
    return decission




x = Menu()

if x == 1:
    writeItems()
elif x == 2:
    updateItems()
elif x == 3:
    deleteItems()
elif x == 4:
    buyProducts()
# createAccount()