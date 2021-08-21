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


def writeItems():
    path = "shopingFile.xlsx"
    wb_obj = openpyxl.load_workbook(path)
    sheet = wb_obj['items']
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

    while(True):
        

        name = input("Product Name: ")
        
        name = name.capitalize()
        decission = True
        if checkExistingProduct(name) == True and decission == True:
            print("Item is already exist in the list!")
            printItems()
            itemCheck = input("Do you want to update the item? y/n ")
            if itemCheck == 'y' or itemCheck == 'Y':
                print()
                updateItems()
            else:
                decission = False
                break
                
        elif decission == True:
            id += 1
            sheet.cell(i,1,value = id)
            lst = giveID()


            sheet.cell(i,2,value = name)
            
            price = float(input("Enter product unit price(BDT): "))
            
            sheet.cell(i,3,value = price)
            quantity = float(input("Quantity: "))
            
            sheet.cell(i,4,value = quantity)
            check = input("Want to add more items? y/n: ")
            i += 1
            print(end = '\n')
            if check == 'n' or check == 'N':
                break
    
    
    wb_obj.save('shopingFile.xlsx') 
    printItems()
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
    while(True):
        id = int(input("Enter your product ID:" ))
        id = (lst.index(id))+2
        print(id)

        name = input("Product Name: ")
        
        sheet.cell(id,2,value = name)
        
        price = float(input("Enter product unit price(BDT): "))
        
        sheet.cell(id,3,value = price)
        quantity = float(input("Quantity: "))
        
        sheet.cell(id,4,value = quantity)
        check = input("Want to add more items? y/n: ")
        if check == 'y' or check == 'Y':
            updateItems() 
        else: break
    
    
    wb_obj.save('shopingFile.xlsx')
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


deleteItems()