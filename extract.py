import tabula
import pandas as pd
import re
import subprocess
import os

def extract(file):
    df = tabula.read_pdf(file, pages="1,2", area=(9.405,41.085,600.435,770.715),nospreadsheet = True,guess=False, pandas_options={'header': None})
    pid = df.iloc[3,0]
    name = df.iloc[4,0].replace(" ","").replace(".","").lower().title()
    
    pid = re.findall(r'\d+',pid)[0]
    pid ="P{}".format(pid)

    # this extracts hours attempted and gpa
    vals = df.iloc[:,0]
    vals = vals[vals.str.contains('TRANSCRIPT|T RANSCRIPT',na=False)].to_string()
    vals = vals.split('|')

    earned = float(vals[2])
    gpa = "Fail" if float(vals[3])<3 else u'\u2714'

    df = df.iloc[:,0]
    df = df[df.str.contains('Fa\d+\s+|Sp\d+\s+|Su\d+\s+',na=False)]

    # this extracts lines that begin with fall, spring or summer untill it gets to the |
    df1 = df[df.str.contains('^Fa\d+\s+|^Sp\d+\s+|^Su\d+\s+',na=False)]
    df1 = df1.iloc[:, ].str.extract('([^\|]+)')

    # this extracts lines where course information is in the middle of the line

    df2 = df[df.str.contains('^[^Fa\d+\s+]|^[^Sp\d+\s+]|^[^Su\d+\s+]',na=False)]
    df2 = df.iloc[:, ].str.extract('(Fa\d+\s+\w*(.*)|Sp\d+\s+\w*(.*)|Su\d+\s+\w*(.*))')
    df2 = df2.iloc[:,0:1]

    total = pd.concat([df1,df2],ignore_index=True)
    total.rename(columns={ total.columns[0]: "course" }, inplace = True)
    total = total.drop_duplicates()
    total['course'] =total['course'].str.split('|').str[0]
    total['course'] =total['course'].str.lstrip()

    total['term'] = total['course'].str[0:4]
    total['class'] = total['course'].str[4:13]
    total['course'] = total['course'].str[13:]
    total['hours'] = total['course'].str.extract("(\d*\.?\d+)", expand=True)
    total['course'] = total['course'].str.replace(r"(\d*\.?)","", regex=True)
    total['course'] = total['course'].str.lstrip()
    total['grade'] = total['course'].str[0:2]
    total['acad_year'] = '20'+total['term'].str[2:]
    total.drop(columns=['course'], inplace=True)
    total.drop_duplicates(inplace=True)
    total.sort_values(by=['acad_year','term'], inplace=True)
  
    fn = "Fail" if total.grade.str.contains('FN',regex=True).any() else u'\u2714'
    pr = 'Fail' if total.grade.str.contains('PR', regex=True).any() else u'\u2714'
    wp = "Fail" if total.grade.str.contains('WP', regex=True).any() else u'\u2714'
    ip = "Fail" if total.grade.str.contains('IP', regex=True).any() else u'\u2714'
    lower_C = "Fail" if total.grade.str.contains( 'C\-|^D$|^F$', regex=True).any() else u'\u2714'
    
    file = os.path.basename(file)

    temp = [name, file, gpa, pr, fn,wp,ip,lower_C]
    return(temp)
    #return(total)
# pdf launcher
def launcher(file):
    command = "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"
    return subprocess.Popen([command, file], stdout=subprocess.PIPE)

if __name__ == "__main__":
    extract()
    launcher()