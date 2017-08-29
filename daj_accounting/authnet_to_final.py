from openpyxl import load_workbook
#credited is being taken out of monthly column in member sheet
#input('type the file name to get a total: ')
#june_auth_gtf.xlsx or july_auth_gtf.xlsx
authnet_file =  'june_auth_gtf.xlsx'
test_gtf_file = 'final_gtf.xlsx'

wb = load_workbook(authnet_file)
wb2 = load_workbook(test_gtf_file)

ws = wb[wb.sheetnames[0]]
ws2 = wb2[wb2.sheetnames[0]]

def get_total_and_list():
    total = 0
    people_list = set([])
    for row in range(1, ws.max_row+1):
        for col in 'Y':
            first_name = ws["{}{}".format(col, row)].value
            cell_name = ws["{}{}".format(col, row)]
        for col in 'Z':
            last_name = ws["{}{}".format(col, row)].value
        for col in 'B':
            cell = ws["{}{}".format(col, row)].value
        for col in 'C':
            value = ws["{}{}".format(col, row)].value
            if cell == 'Credited':
                total -= float(value)
                new_person = Person(first_name, last_name, -value, cell_name)
                people_list.add(new_person)
                #print('this value should be negative', ws[cell_value].value) 
            if cell == 'Settled Successfully':
                total += float(value)
                new_person = Person(first_name, last_name, value, cell_name)
                people_list.add(new_person)
    print('TOTAL: {0}'.format(total))
    return people_list

class Person:
    def __init__(self, first_name, last_name, amount = 0, cell = None):
        if first_name == None:
            first_name = 'none'
            #print(cell)
        if last_name == None:
            last_name = 'none'
            #print(cell)

        self.first_name = first_name.lower()
        self.last_name = last_name.lower()
        self.amount = amount
        self.cell = cell
        
peeps = get_total_and_list()
#print('people_list is ', len(peeps), ' long')

def map_to_final():
    members_list = set([])
    for row in range(2, ws2.max_row+1):
        for col in 'G':
            initial = "{}{}".format(col, row)
        for col in 'H':
            monthly = "{}{}".format(col, row)
        for col in 'I':
            hundreds = "{}{}".format(col, row)
        for col in 'A':
            f_name = ws2["{}{}".format(col, row)].value.lower()
        for col in 'B':
            l_name = ws2["{}{}".format(col, row)].value.lower()  
            for person in peeps:
                if person.first_name == f_name and person.last_name == l_name:      
                    new_person = Person(f_name, l_name)
                    members_list.add(new_person)                    
                    if person.amount >= 1000:
                        if isinstance(ws2[initial].value, int) or isinstance(ws2[initial].value, float): 
                            ws2[initial] = float(ws2[initial].value) + person.amount
                        else:
                            ws2[initial] = person.amount                
                        #wb2.save('finished2_june_gtf.xlsx')
                    elif person.amount == 100 or person.amount == 200:
                        if isinstance(ws2[hundreds].value, int) or isinstance(ws2[hundreds].value, float): 
                            ws2[hundreds] = float(ws2[hundreds].value) + person.amount
                        else:
                            ws2[hundreds] = person.amount
                        #wb2.save('finished2_june_gtf.xlsx')
                    elif (person.amount != 100 or person.amount != 200) and person.amount < 1000:
                        if isinstance(ws2[monthly].value, int) or isinstance(ws2[monthly].value, float): 
                            ws2[monthly] = float(ws2[monthly].value) + person.amount
                        else:
                            ws2[monthly] = person.amount        
                        #wb2.save('finished2_june_gtf.xlsx')
    return members_list

members_list = map_to_final()                                
members_names = {member.first_name + ' ' + member.last_name for member in members_list}
peeps_names = {person.first_name + ' ' + person.last_name for person in peeps}

def test_bool():
    i = ws2.max_row
    j = i + 1
    new_peeps = peeps_names.difference(members_names)    
    for person in peeps:
#        print(person.first_name + ' ' + person.last_name)
        for name in new_peeps:
#            print(name)
            if person.first_name + ' ' + person.last_name == name:
                for row in range(794, 1000):
                    for col in 'G':
                        initial = "{}{}".format(col, row)
                    for col in 'H':
                        monthly = "{}{}".format(col, row)
                    for col in 'I':
                        hundreds = "{}{}".format(col, row)
                    for col in 'A':
                        f_name = ws2["{}{}".format(col, row)].value.lower()
                    for col in 'B':
                        l_name = ws2["{}{}".format(col, row)].value.lower()
                        if person.amount >= 1000:
                            print('name ', name, '== ', person.first_name, ' ', person.last_name)
                            print('f_name and l_name ', f_name + ' ' + l_name)
                            if isinstance(ws2[initial].value, int) or isinstance(ws2[initial].value, float): 
                                if person.first_name == f_name and person.last_name == l_name:
                                    ws2[initial] = float(ws2[initial].value) + person.amount
                                    print(ws2[initial])
                            else:
                                print(ws2["{}{}".format('A', ws2.max_row + 1)])
                                ws2["A" + str(ws2.max_row + 1)] = person.first_name
                                ws2["B" + str(ws2.max_row)] = person.last_name
                                ws2["G" + str(ws2.max_row)] = person.amount
                                j += 1
#                            wb2.save('finished2_june_gtf.xlsx')
                        elif person.amount == 100 or person.amount == 200:
                            if isinstance(ws2[hundreds].value, int) or isinstance(ws2[hundreds].value, float):
                                if person.first_name == f_name and person.last_name == l_name:
                                    ws2[hundreds] = float(ws2[hundreds].value) + person.amount
                                    print(ws2[hundreds])
                            else:
                                print(ws2["{}{}".format('A', ws2.max_row)])
                                ws2["A" + str(ws2.max_row + 1)] = person.first_name
                                ws2["B" + str(ws2.max_row)] = person.last_name
                                ws2[hundreds] = person.amount
                                j += 1
#                            wb2.save('finished2_june_gtf.xlsx')
                        elif (person.amount != 100 or person.amount != 200) and person.amount < 1000:
                            if isinstance(ws2[monthly].value, int) or isinstance(ws2[monthly].value, float):
                                if person.first_name == f_name and person.last_name == l_name:
                                    ws2[monthly] = float(ws2[monthly].value) + person.amount
                                    print('total added to monthly ', ws2[monthly])
                            else:
                                print(ws2["{}{}".format('A', ws2.max_row)])
                                ws2["A" + str(ws2.max_row + 1)] = person.first_name
                                ws2["B" + str(ws2.max_row)] = person.last_name
                                ws2[monthly] = person.amount
                                j += 1
#                            wb2.save('finished2_june_gtf.xlsx')
    print('FINISHED')                        
#    return new_peeps
#final step: add in new members info from peeps list    
#print(test_bool())
test_set = test_bool()
#print(peeps[5].last_name)
#print('test_peeps: ', len(test_set))
print('new_peeps: ', len(members_list))
print('peeps: ', len(peeps))
