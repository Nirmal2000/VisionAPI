import json
import math
from openpyxl import Workbook
import os


column_dict = ['registrationno.B','regn.noB',
    'reg.dt:C','regdt:C'
    'ch.noF','chassisno.F','chno',
    'enoE','engineE','engineno.E',
    'nameA','name.A','owner\'snameA','name&addressA',
    'mfg.dt.D','monthandyearofmfg.D','month/yrofD','mfgdt']

def distance(x1 , y1 , x2 , y2): 
    return math.sqrt(math.pow(x2 - x1, 2) +
                math.pow(y2 - y1, 2) * 1.0) 


def lineFromPoints(P,Q): 
    a = Q[1] - P[1] 
    b = P[0] - Q[0]  
    c = -(a*(P[0]) + b*(P[1]) )    
    return a,b,c

def shortest_distance(x1, y1, a, b, c):  
       
    d = abs((a * x1 + b * y1 + c)) / (math.sqrt(a * a + b * b)) 
    return d


def get_line(P,Q,word_size,text_annotations,word_set):
    
    lines = []
    a,b,c = lineFromPoints(P,Q)
    for word_desc in text_annotations:
        try:
            x1 = word_desc['bounding_poly']['vertices'][0]['x']
            y1 = word_desc['bounding_poly']['vertices'][0]['y']
            x2 = word_desc['bounding_poly']['vertices'][3]['x']
            y2 = word_desc['bounding_poly']['vertices'][3]['y']
            x3 = word_desc['bounding_poly']['vertices'][1]['x']
            x2 = (x1+x2)/2
            y2 = (y1+y2)/2
            if shortest_distance(x2,y2,a,b,c)<=(word_size):
                word_set.add(word_desc['description'].lower()+str(x1)+str(y1))
                lines.append([x1,word_desc['description'],x3])
        except:
            pass


    lines = sorted(lines, key=lambda x: x[0])
    line=[]
    for i,dei in enumerate(lines):
        try:
            line.append([dei[1],lines[i+1][0]-dei[2]])
        except:
            line.append([dei[1],0])
    return line



def parse_it(line,i,cell_col):
    
    if line[i][0]==':':
        i=i+1
    if(cell_col=='C' or cell_col=='D'):
        return line[i][0]
    retstr =''
    while True:
        retstr+=line[i][0]+' '    
        if i+1>=len(line) or line[i][1]>8:
            break
        
        i+=1
    return retstr


def parse_month(line,i):

    while(line[i][0][0].isdigit()==False):
        print(line[i][0])
        i+=1
    return line[i][0]


def check_match(line,i,col_name):
    cell_col = col_name[-1]
    col_name = col_name[:-1]
    for it in range(i,len(line)):
        if(col_name == line[it][0].lower()):
                    
            return parse_it(line,i+1,cell_col)
        elif it+1<len(line) and  col_name == (line[it][0] + line[it+1][0]).lower():
            
            return parse_it(line,i+2,cell_col)
        elif it+2<len(line) and  col_name == (line[it][0] + line[it+1][0] + line[it+2][0]).lower():
            
            return parse_it(line,i+3,cell_col)
        elif line[it][0].lower() == 'month' and col_name.startswith('month'):            
            return parse_month(line,i+1)
    return '-1'

def save_to_sheet():
    book = Workbook()
    sheet  = book.active
    row_number=0
    for filename in os.listdir('./jsons/'):
        row_number+=1
        print(filename)
        path = './jsons/{}'.format(filename)
        with open(path,'r') as f:
            img_dict = json.load(f)

        word_set = set()
        text_annotations = img_dict['text_annotations'][1:]

        i=0
        lines=[]
        for word_desc in text_annotations:    ## Find the lines     
            try:
                x1 = word_desc['bounding_poly']['vertices'][0]['x']
                y1 = word_desc['bounding_poly']['vertices'][0]['y']        
                if word_desc['description'].lower()+str(x1)+str(y1) not in word_set:              
                        x2 = word_desc['bounding_poly']['vertices'][1]['x']
                        y2 = word_desc['bounding_poly']['vertices'][1]['y'] 
                        x3 = word_desc['bounding_poly']['vertices'][3]['x'] 
                        y3 = word_desc['bounding_poly']['vertices'][3]['y']
                        x4 = word_desc['bounding_poly']['vertices'][2]['x'] 
                        y4 = word_desc['bounding_poly']['vertices'][2]['y'] 
                        x1 = (x1+x3)/2
                        y1 = (y1+y3)/2
                        x2 = (x2+x4)/2
                        y2 = (y2+y4)/2
                        word_size = distance(x1,y1,x3,y3)
                        lines.append(get_line([x1,y1],[x2,y2],word_size,text_annotations,word_set))
                        
            except:
                pass

        
        for line in lines: #Saving to excel
            for i,word in enumerate(line):        
                for col_name in column_dict:
                    cell_insert = check_match(line,i,col_name)
                    if(cell_insert == '-1'):
                        continue
                    else:                    
                        ins_col = ord(col_name[-1]) - ord('A') + 1                    
                        if sheet.cell(row=row_number,column=ins_col).value == None:
                            sheet['{}{}'.format(col_name[-1],row_number)] = cell_insert
                        break
        
        
        sheet['G{}'.format(row_number)] = str(filename)
                    



    book.save('license.xlsx')