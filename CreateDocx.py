import win32api
import win32print
import tempfile
import time
import os
from docxtpl import DocxTemplate, RichText


def blank(length,str):
    strlen = len(str)
    blank = int((length-strlen)/2)
    bstr = str
    while blank > 0:
        bstr = " "+bstr
        bstr = bstr+" "
        blank -= 1
    return bstr


def createdocx(srodtime,sproddate,sdestination,
            smodel,sr,spdo,slot,stotalqty,
            spcs,sqtypallethead,sqtypallet,spallet,srecdate,sremark):
    root_dir = os.path.abspath('.')
    tpl = DocxTemplate(root_dir + '/config/simple.docx')
    rt = RichText('\n')
    rt.add('\n')
    RODTIME = RichText(blank(17,srodtime))  #19-2
    DESTINATION = RichText(sdestination)
    PRODDATE = RichText(blank(19,sproddate)) #21-2
    MODEL = RichText(blank(18,smodel)) #18-2
    R = RichText(blank(18,sr))#20-2
    PDO = RichText(blank(18,spdo))#20-2
    LOT = RichText(blank(20,slot))#22-2
    TOTALQTY = RichText(blank(30,stotalqty))#32-2
    PCS = RichText(blank(20,spcs))#22-2
    QTYPALLET = RichText(blank(54,sqtypallethead+sqtypallet))#56-2
    PALLET = RichText(blank(15,spallet))#17-2
    RECDATE = RichText(blank(24,srecdate))#26-2
    REMARK = RichText(blank(76,sremark))#78-2
    
    if int(stotalqty)>int(sqtypallet):
        cpallet=int(stotalqty)/int(sqtypallet)
        
        box=[]
        count=1
        while count<=cpallet:
            dict = { 
                    'blank': rt,
                    'destination':DESTINATION,
                    'RODTIME': RODTIME,
                    'PRODDATE': PRODDATE,
                    'MODEL': MODEL,
                    'R': R,
                    'PDO': PDO,
                    'LOT': LOT,
                    'TOTALQTY': TOTALQTY,
                    'PCS': PCS,
                    'QTYPALLET': QTYPALLET,
                    'PALLET': RichText(blank(15,str(count))),
                    'RECDATE': RECDATE,
                    'REMARK':REMARK
                    }
            box.append(dict)
            count += 1
        if int(stotalqty)%int(sqtypallet)!=0:
            dict = { 
                    'blank': '',
                    'destination':DESTINATION,
                    'RODTIME': RODTIME,
                    'PRODDATE': PRODDATE,
                    'MODEL': MODEL,
                    'R': R,
                    'PDO': PDO,
                    'LOT': LOT,
                    'TOTALQTY': TOTALQTY,
                    'PCS': PCS,
                    'QTYPALLET': RichText(blank(54,str(int(stotalqty)-int(sqtypallet)*(count-1)))),
                    'PALLET': RichText(blank(15,str(count))),
                    'RECDATE': RECDATE,
                    'REMARK':REMARK
                    }
            box.append(dict)
        context = {
            'box': box
        }        
        tpl.render(context)
        if not os.path.exists(root_dir + "/file"):
            os.makedirs(root_dir + "/file")
        tpl.save(time.strftime(root_dir +"/file/%Y%m%d_%H%M%S", time.localtime())+'.docx')
    elif int(stotalqty)<=int(sqtypallet):
        box=[]
        dict = { 
                    'blank': '',
                    'destination':DESTINATION,
                    'RODTIME': RODTIME,
                    'PRODDATE': PRODDATE,
                    'MODEL': MODEL,
                    'R': R,
                    'PDO': PDO,
                    'LOT': LOT,
                    'TOTALQTY': TOTALQTY,
                    'PCS': PCS,
                    'QTYPALLET': TOTALQTY,
                    'PALLET': RichText(blank(15,str(1))),
                    'RECDATE': RECDATE,
                    'REMARK':REMARK
                    }
        box.append(dict)
        context = {
            'box': box
        }
        tpl.render(context)
        if not os.path.exists(root_dir + "/file"):
            os.makedirs(root_dir + "/file")
        tpl.save(time.strftime(root_dir + "/file/%Y%m%d_%H%M%S", time.localtime()) + '.docx')

def printdocx(srodtime,sproddate,sdestination,
            smodel,sr,spdo,slot,stotalqty,
            spcs,sqtypallethead,sqtypallet,spallet,srecdate,sremark):
    root_dir = os.path.abspath('.')
    tpl = DocxTemplate(root_dir + '/config/simple.docx')
    rt = RichText('\n')
    rt.add('\n')
    RODTIME = RichText(blank(17,srodtime))  #19-2
    DESTINATION = RichText(sdestination)
    PRODDATE = RichText(blank(19,sproddate)) #21-2
    MODEL = RichText(blank(18,smodel)) #18-2
    R = RichText(blank(18,sr))#20-2
    PDO = RichText(blank(18,spdo))#20-2
    LOT = RichText(blank(20,slot))#22-2
    TOTALQTY = RichText(blank(30,stotalqty))#32-2
    PCS = RichText(blank(20,spcs))#22-2
    QTYPALLET = RichText(blank(54,sqtypallethead+sqtypallet))#56-2
    PALLET = RichText(blank(15,spallet))#17-2
    RECDATE = RichText(blank(24,srecdate))#26-2
    REMARK = RichText(blank(76,sremark))#78-2
    
    if int(stotalqty)>int(sqtypallet):
        cpallet=int(stotalqty)/int(sqtypallet)
        
        box=[]
        count=1
        while count<=cpallet:
            dict = { 
                    'blank': rt,
                    'destination':DESTINATION,
                    'RODTIME': RODTIME,
                    'PRODDATE': PRODDATE,
                    'MODEL': MODEL,
                    'R': R,
                    'PDO': PDO,
                    'LOT': LOT,
                    'TOTALQTY': TOTALQTY,
                    'PCS': PCS,
                    'QTYPALLET': QTYPALLET,
                    'PALLET': RichText(blank(15,str(count))),
                    'RECDATE': RECDATE,
                    'REMARK':REMARK
                    }
            box.append(dict)
            count += 1
        if int(stotalqty)%int(sqtypallet)!=0:
            dict = { 
                    'blank': '',
                    'destination':DESTINATION,
                    'RODTIME': RODTIME,
                    'PRODDATE': PRODDATE,
                    'MODEL': MODEL,
                    'R': R,
                    'PDO': PDO,
                    'LOT': LOT,
                    'TOTALQTY': TOTALQTY,
                    'PCS': PCS,
                    'QTYPALLET': RichText(blank(54,str(int(stotalqty)-int(sqtypallet)*(count-1)))),
                    'PALLET': RichText(blank(15,str(count))),
                    'RECDATE': RECDATE,
                    'REMARK':REMARK
                    }
            box.append(dict)
        context = {
            'box': box
        }        
        tpl.render(context)
        filename=tempfile.mktemp(".docx")
        tpl.save(filename)
        win32api.ShellExecute (
          0,
          "print",
          filename,
          # If this is None, the default printer will
          # be used anyway.
          '/d:"%s"' % win32print.GetDefaultPrinter (),
          ".",
          0
        )
    elif int(stotalqty)<=int(sqtypallet):
        box=[]
        dict = { 
                    'blank': '',
                    'destination': DESTINATION,
                    'RODTIME': RODTIME,
                    'PRODDATE': PRODDATE,
                    'MODEL': MODEL,
                    'R': R,
                    'PDO': PDO,
                    'LOT': LOT,
                    'TOTALQTY': TOTALQTY,
                    'PCS': PCS,
                    'QTYPALLET': TOTALQTY,
                    'PALLET': RichText(blank(15,str(1))),
                    'RECDATE': RECDATE,
                    'REMARK':REMARK
                    }
        box.append(dict)
        context = {
            'box': box
        }        
        tpl.render(context)
        filename=tempfile.mktemp(".docx")
        tpl.save(filename)
        win32api.ShellExecute (
          0,
          "print",
          filename,
          # If this is None, the default printer will
          # be used anyway.
          '/d:"%s"' % win32print.GetDefaultPrinter (),
          ".",
          0
        )

if __name__=='__main__':
    createdocx("2019-09-01","2019-09-01","法国",
            "XF8505","996A-56-10","550750","13","414",
            "3","138箱*3=","414","2","2019.7.29","5550.20519RM")
    #RODTIME="2019-09-01"  #19-2
    #PRODDATE="2019-09-01" #21-2
    #MODEL="XF8505" #18-2
    #R="996A-56-10"#20-2
    #PDO="550750"#20-2
    #LOT="13"#22-2
    #TOTALQTY="1002"#32-2
    #PCS="3"#22-2
    #QTYPALLET="138*33=414"#41-2
    #PALLET="2"#17-2
    #RECDATE="2019.7.29"#26-2
    #REMARK="5550.20519RM"#78-2
 