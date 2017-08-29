from openpyxl import load_workbook

#input('type the file name to get a total: ')
#june_auth_gtf.xlsx or july_auth_gtf.xlsx
authnet_file =  'june_auth_gtf.xlsx'
test_gtf_file = 'final_gtf.xlsx'

wb = load_workbook(authnet_file)
wb2 = load_workbook(test_gtf_file)

ws = wb[wb.sheetnames[0]]
ws2 = wb2[wb2.sheetnames[0]]

def get_total():
    total = 0
    for row in range(1, ws.max_row+1):
        for col in 'B':
            cell = ws["{}{}".format(col, row)].value
        for col in 'C':
            value = ws["{}{}".format(col, row)].value
            if cell == 'Credited':
                total -= float(value)
                #print('this value should be negative', ws[cell_value].value) 
            if cell == 'Settled Successfully':
                total += float(value)
    print('TOTAL: {0}'.format(total))

get_total()    

def peeps_all():
    print('Building Members List...')
    peeps_set = set([])
    for row in range(1, ws.max_row+1):
        for col in 'Y':
            f_name = ws["{}{}".format(col, row)].value
#            cell_name = ws["{}{}".format(col, row)]
        for col in 'Z':
            l_name = ws["{}{}".format(col, row)].value
        for col in 'B':
            cell = ws["{}{}".format(col, row)].value
            if cell == 'Credited' or cell == 'Settled Successfully':
                new_person = Person(f_name, l_name)
                peeps_set.add(new_person)
#                print('new person added ', f_name, l_name)
    return peeps_set

#if person is the issue-- build a list of names-- get unique names
#for names in list, build a person



class Person:
    def __init__(self, first_name, last_name, initial = 0, monthly = 0, \
                                               hundreds = 0, cell = None):
        if first_name == None:
            first_name = 'none'
            #print(cell)
        if last_name == None:
            last_name = 'none'
            #print(cell)

        self.first_name = first_name.lower()
        self.last_name = last_name.lower()
        self.initial = initial
        self.monthly = monthly
        self.hundreds = hundreds
        self.cell = cell

peeps_names = {person.first_name + ' ' + person.last_name for person in peeps_all()}
#print('all peeps ', len(peeps_all()))

def get_amounts():
    print('Adding Transactions To Members...')
    peeps = peeps_all()
    
    for row in range(1, ws.max_row+1):
        for col in 'Y':
            if ws["{}{}".format(col, row)].value == None:
                ws["{}{}".format(col, row)] = 'none'
            else:
                f_name = ws["{}{}".format(col, row)].value.lower()
#                cell_name = ws["{}{}".format(col, row)]
        for col in 'Z':
            if ws["{}{}".format(col, row)].value == None:
                ws["{}{}".format(col, row)] = 'none'
            else:
                l_name = ws["{}{}".format(col, row)].value.lower()
        for col in 'B':
            cell = ws["{}{}".format(col, row)].value
        for col in 'C':
            value = ws["{}{}".format(col, row)].value
            
            if cell == 'Settled Successfully':
                for person in peeps:
                    if person.first_name == f_name and person.last_name == l_name:
                        if value >= 1000:
                            person.initial += value
                            print(person.first_name, person.last_name, person.initial)
                        elif value == 100 or value == 200:
                            person.hundreds += value                        
                        elif (value != 100 or value != 200) and value < 1000:
                            person.monthly += value
            elif cell == 'Credited':
                for person in peeps:
                    if person.first_name == f_name and person.last_name == l_name:
                        if value >= 1000:
                            person.initial -= value
                        elif value == 100 or value == 200:
                            person.hundreds -= value                        
                        elif (value != 100 or value != 200) and value < 1000:
                            person.monthly -= value
    #                print(person.first_name, ' ', person.last_name, ' ', person.initial)
    return peeps                        

peeps = get_amounts()

def map_to_final():
    members_set = set([])
    
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
                    #create a list of current members--to be used later#
                    new_person = Person(f_name, l_name)
                    members_set.add(new_person)
                    #populate cells with data#
                    ws2[initial] = person.initial
                    ws2[monthly] = person.monthly
                    ws2[hundreds] = person.hundreds
    print('Finished Updating...')
    return members_set

members_set = map_to_final()

def add_new_members():
    print('Adding New Members...')
    #compare sets for difference#    
    members_names = {member.first_name + ' ' + member.last_name for member in members_set}
    peeps_names = {person.first_name + ' ' + person.last_name for person in peeps}
    print('peeps ', len(peeps_names))    
    print('members ', len(members_names))    
    new_peeps = peeps_names.difference(members_names)
    print('new_peeps ', len(new_peeps))
    for person in peeps:
        for name in new_peeps:#the assumption is there is only one person in peeps
            if person.first_name + ' ' + person.last_name == name:
#                print('person ', person.first_name + ' ' + person.last_name)
#                print('name ', name)
                ws2["A" + str(ws2.max_row + 1)] = person.first_name
                ws2["B" + str(ws2.max_row)] = person.last_name
                ws2["G" + str(ws2.max_row)] = person.initial
                ws2["H" + str(ws2.max_row)] = person.monthly
                ws2["I" + str(ws2.max_row)] = person.hundreds
#                wb2.save('finished2_june_gtf.xlsx')
    print('FINISHED')
add_new_members()
