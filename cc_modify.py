import xls_read

class cc_modify:
    def __init__(self, line_in_xls, workbook):
        self.line = workbook.getLine(line_in_xls)
        print(self.line)
        with open(self.line[0], 'r') as file:
            self.lines = file.readlines()
            print(self.lines[2])
        
        if(self.line[2] == 'ADD'):
            if(self.line[3] == 'Below'):
                for i,row in enumerate(self.lines):                 
                    if(self.line[4] in row):
                        print(self.line[5])
                        self.lines[i-1] = self.lines[i-1] + '\n' +self.line[5] + '\n' 
                        
            if(self.line[3] == 'Above'):
                for i,row in enumerate(self.lines):
                    if(self.line[4] in row):
                        print(self.line[5])
                        self.lines[i-1] = self.line[5] + '\n' + self.lines[i-1] + '\n' 
                        
        if(self.line[2] == 'Modify' ):
            start_row_number = -1
            stop_row_number = -1
            print('Afisare code change ' + str(self.line[6]))
            for i,row in enumerate(self.lines):
                if( start_row_number == -1):
                    if(self.line[7] in row):
                        print('incepere modificare -----------------------------' + '\n')
                        print('/*' + '\n' + self.lines[i-1])
                        self.lines[i-1] = '/*' + '\n' + self.lines[i-1]                      
                        start_row_number = i;
                        print(self.line[8])
                else:
                    if(stop_row_number == -1):
                        if(self.line[8] in row):
                            print(self.lines[i] + '\n' + '*/' + '\n' + self.line[6] + '\n')
                            self.lines[i] = self.lines[i] + '\n' + '*/' + '\n' + self.line[6] + '\n'
                            stop_row_number = i
                            print('sfarsit modificare -----------------------------' + '\n')
                            print('start-row =' + str(start_row_number) + '   stop-row = ' + str(stop_row_number))
                    else:
                        break
                
        with open(self.line[0], 'w') as file:
            file.writelines(self.lines)

def main():
    file = xls_read.ModuleWorkbook("Memstack_CodeChanges.xlsx")
    CC = cc_modify(3, file)
    
    #line = file.getLine(3)
    #print(line)
    
if __name__ == '__main__':
    main()