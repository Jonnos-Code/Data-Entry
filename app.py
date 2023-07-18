import PySimpleGUI as sg
import keyboard as kb
from time import sleep
from functools import partial
import cv2
import io
import requests
import json
from thefuzz import fuzz,process
import numpy as np
from PIL import ImageGrab
import win32gui

sg.LOOK_AND_FEEL_TABLE['DarkBlack1']['INPUT'] = '#e4dcaF'
sg.theme('DarkBlack1')

Signatures = {
    'Atacamite':1800,
    'Felsic':1770,
    'Gneiss':1840,
    'Igneous':1950,
    'Obsidian':1790,
    'Shale':1730,
    'Quartzite':1820,
    'C Type':1700,
    'E Type':1900,
    'M Type':1850,
    'P Type':1750,
    'Q Type':1870,
    'S Type':1720
}
ElementNames = {
    'Quantanium (RAW)':['Quant.',682,'242676250'],
    'Gold (ORE)':['Gold',644,'1559243841'],
    'Taranite (RAW)':['Taranite',340,'1629693614'],
    'Bexalite (RAW)':['Bexalite',231,'265944129'],
    'Borase (ORE)':['Borase',150,'753994680'],
    'Laranite (RAW)':['Laranite',384,'835894443'],
    'Hephaestanite (RAW)':['Heph.',107,'2125864444'],
    'Agricium (ORE':['Agricium',240,'853527979'],
    'Beryl (RAW)':['Beryl',92,'1930776786'],
    'Titanium (ORE)':['Tita.',150,'1027926027'],
    'Tunsten (ORE)':['Tungst.',644,'762911649'],
    'Quartz (RAW)':['Quartz',89,'2140254015'],
    'Corundum (RAW)':['Corund.',134,'1355904006'],
    'Iron':['Iron',263,'403009114'],
    'Copper (ORE)':['Copper',299,'268255631'],
    'Aluminum (ORE)':['Alum.',91,'788137754'],
    'Inert Materials':['Inert',33]
    }
locations = {
    # Crusader not needed; no mining
    'Daymar':['Wolf Point Aid Shelter','Nuen Waste Management','Dunlow Ridge Aid Shelter','NT-999-XVI','TPF','The Garden','Security Post Prashad',"Brio's Breaker Yard",'Security Post Thaquray','Security Post Moluto','Kudre Ore Mine (Closed)','Shubin Mining Facility SCD-1','Eager Flats Aid Shelter','Kudre Ore','Tamdon Plains Aid Shelter','Bountiful Harvest Hydroponics','ArcCorp Mining Area 141'],
    'Cellin':['Ashburn Channel Aid Shelter','PRIVATE PROPERTY','Security Post Dipur','Security Post Criska','NT-999-XV','Terra Mills HydroFarm','Hickes Research Outpost','Julep Ravine Aid Shelter','Security Post Lespin','Tram & Myers Mining','Mogote Aid Shelter',"Flanagan's Ravine Aid Shelter",'Gallete Family Farms'],
    'Yela':['Talarine Divide Aid Shelter','Security Post Opal','Aston Ridge Aid Shelter','NT-999-XXII','Utopia','ArcCorp Mining Area 157','Security Post Wan','Nakamura Valley Aid Shelter','Deakins Research Outpost','Benson Mining Outpost','NT-999-XX','Afterlife','Kosso Basin Aid Shelter',"Connor's"],
    'Hurston':['Lorville','HDMS-Edmond','HDES-Calthrope (NA)','Reclamation & Disposal Orinth','HDSF-Sherman','HDMS-Pinewood','HDMS-Hadley','HDSF-Colfax','HDSF-Tompkins','HDSF-Tamar','HDMS-Thedus','HDMS-Stanhope','HDSF-Hiram','HDMS-Oparei','HDSF-Millerand','HDSF-Barnabas'],
    'Arial':['HDMS-Bezdek','HDMS-Lathan'],
    'Aberdeen':['HDMS-Norgaard','HDMS-Anderson','HDES-Dobbs (NA)','Klescher Rehabilitation Facility','Barton Flats Aid Shelter'],
    'Magda':['HDMS-Hahn','HDMS-Perlman'],
    'Ita':['HDMS-Ryder','HDMS-Woodruff'],
    # ArcCorp not needed; no mining
    'Lyria':['"Wheeler\'s"','Humboldt Mines','Buckets','Shubin Processing Facility SPAL-9','Shubin Processing Facility SPAL-3','Launch Pad','Elsewhere',"Teddy's Playhouse",'Shubin Processing Facility SPAL-7','Shubin Processing Facility SPAL-12','Operations Depot Lyria-1','Loveridge Mineral Reserve','Shubin Mining Facility SAL-2','Shubin Processing Facility SPAL-16','The Orphanage','Shubin Processing Facility SPAL-21','Shubin Mining Facility SAL-5','"The Pit"'],
    'Wala':['ArcCorp Mining Area 061','ArcCorp Mining Area 048','ArcCorp Mining Area 045','ArcCorp Processing Center 115','ArcCorp Processing Center 123',"Samson & Son's Salvage Center",'Shady Glen Farms','Good Times Temple','ArcCorp Mining Area 056','Lost and Found'],
    'MicroTech':['New Babbage','Shubin Mining Facility SMO-10','Shubin Mining Facility SMO-13','Shubin Mining Facility SMO-22','The Necropolis','Ghost Hollow','Outpost 54','Shubin Mining Facility SMO-18','Calhoun Pass Emergency Shelter','Nuiqsut Emergency Shelter','Point Wain Emergency Shelter','Clear View Emergency Shelter','Rayari Deltana Research Outpost','MT DataCenter TMG-XEV-2','MT Datacenter KH3-AAE-L','MT DataCenter 8FK-Q2K-K','MT DataCenter QVX-J88-J','MT DataCenter D79-ECG-R','MT DataCenter 5WQ-R2V-C','MT DataCenter E2Q-NSG-Y','MT DataCenter 2UB-RB9-5','MT DataCenter L8P-JUC-8 (Offline)'],
    'Calliope':['Shubin Processing Facility SPMC-5','Shubin Processing Facility SPMC-10','Shubin Processing Facility SPMC-3','Shubin Processing Facility SPMC-14','Shubin Processing Facility SPMC-11','Shubin Mining Facility SMCa-6','Shubin Mining Facility SMCa-8','Rayari Kaltag Research Outpost','Rayari Anvik Research Outpost'],
    'Clio':['Rayari Cantwell Research Outpost','Rayari McGrath Research Outpost'],
    'Euterpe':["Bud's Growery",'The Icebreaker','Devlin Scrap & Salvage']
}
spaceloc = ['Yela Belt','Aaron Halo',
            'ARC L1-A','ARC L2-A','ARC L3-A','ARC L4-A','ARC L5-A',
            'CRU L1-A','CRU L2-A','CRU L2-B','CRU L2-C','CRU L3-A','CRU L4-A','CRU L5-A',
            'HUR L1-A','HUR L2-A','HUR L3-A','HUR L4-A','HUR L5-A',
            'MIC L1-A','MIC L1-B','MIC L2-A','MIC L2-B','MIC L2-C','MIC L3-A','MIC L4-A','MIC L5-A'
]
regions = ['use1c','euw1a','euw1b','euw1c','ape1a','apse2a']

def ScreenCap(name):
    toplist, winlist = [], []
    def enum_cb(hwnd,results):
        winlist.append((hwnd, win32gui.GetWindowText(hwnd)))
    win32gui.EnumWindows(enum_cb,winlist)

    winob = [(hwnd, title) for hwnd, title in winlist if name.lower() in title.lower()]
    hwnd = winob[0][0]

    win32gui.SetForegroundWindow(hwnd)
    bbox = win32gui.GetWindowRect(hwnd)
    return ImageGrab.grab(bbox,all_screens=True)
def OCR(image):
    setup = {
        'isOverlayRequired':False,
        'apikey':'K88799935788957',
        'language':'eng',
        'OCREngine':2,
        'scale':False,
        'filetype':'PNG'
        }
    buffer = io.BytesIO(cv2.imencode('.png',image)[1])
    output = requests.post('https://api.ocr.space/parse/image',files={'image': buffer},data=setup).content.decode()
    return [x['LineText'] for x in json.loads(output)['ParsedResults'][0]['TextOverlay']['Lines']]
def VMImage(image):
    image = np.concatenate((image[int(image.shape[0]/3):int(1.5*image.shape[0]/3),int(image.shape[1]/2-image.shape[0]/1.4):int(image.shape[1]/2-image.shape[0]/2.5)],
                            image[int(image.shape[0]/3):int(1.6*image.shape[0]/3),int(image.shape[1]/2+image.shape[0]/2.5):int(np.ceil(image.shape[1]/2+image.shape[0]/1.4))]),axis=0)
    output = OCR(image)
    print(output)

    valid, mvalid = False,False
    for i in output:
      if len(i) > 4 and fuzz.partial_ratio(i[:5],'Mass:') > 50:
        mass = int(float(i[i.index(' ')+1:i.index('.')]))
        mvalid = True
      elif fuzz.ratio(i.lower(),'CARRYING'.lower()) > 50:
        comp = output[output.index(i)+1:]
        if len(comp) > 1: valid = True
        break

    if not valid and not mvalid: raise Exception('Unable to read whole scan')
    elif valid and not mvalid: raise Exception('Unable to read left scan')
    elif not valid: raise Exception('Unable to read right scan')

    x=0
    for i in range(len(comp)):
      if fuzz.partial_ratio(comp[i][:11],'Composition') > 10 and '.' in comp[i]:
        comp[i] = (process.extractOne(comp[i][comp[i].index(' '):-9],list(ElementNames.keys()))[0],comp[i][-9:][comp[i][-9:].index('.')-2:comp[i][-9:].index('.')+3]+'%')
        x=i
      else: break
    if x < len(comp): del(comp[x+1:])
    if len(comp) <= 1: raise Exception('Unable to read right scan')
    return mass,comp
def MMImage(image):
    image = image[int(image.shape[0]/3):int(2*image.shape[0]/3),int(image.shape[1]/2+0.33*image.shape[0]):int(image.shape[1]/2+0.77*image.shape[0]),:]
    output = OCR(image)
    print(output)

    valid,mvalid = False,False
    for i in range(len(output)):
        text = output[i]
        if (len(text) >= 4 and text[:4].lower() == 'mass') or (len(text) >= 5 and text[:5].lower() == 'mass:'):
            mvalid = True
            if ' ' in text:
                mass = int(text[text.index(' ')+1:])
            else:
                try:
                    mass = int(output[i+1])
                except:
                    try:
                        mass = int(output[i+3])
                    except:
                        mass = int(output[i+4])
        elif fuzz.ratio(text.lower(),'composition') >= 50:
            comp = output[i+1:]
            valid = True
            break

    if not valid or not mvalid: raise Exception('Unable to read image- try taking a cleaner shot')

    skip,p,mode,s = False,[],False,0
    for i in range(len(comp)):
        if skip:
            if comp[i][0].isnumeric() or comp[i][1].isnumeric():
                mode = True
                continue
            skip = False
            comp[i-1] = (process.extractOne(comp[i],list(ElementNames.keys()))[0],comp[i-1])
            p.append(i)
        elif comp[i][0].isnumeric():
            if mode and not comp[i][0].isnumeric() and not comp[i][1].isnumeric():
                s = i if s==0 else s
                comp[i-s] = (process.extractOne(comp[i],list(ElementNames.keys()))[0],comp[i-s])
                p.append(i)
            elif ' ' in comp[i]:
                comp[i] = (process.extractOne(comp[i][comp[i].index(' ')+1:],list(ElementNames.keys()))[0],comp[i][:comp[i].index(' ')])
                if not comp[i][1][0].isnumeric(): comp[i][1] = comp[i][1][1:]
            elif comp[i][0].isnumeric() or comp[i][1].isnumeric():
                skip = True
                if not comp[i][0].isnumeric(): comp[i] = comp[i][1:]
            else:
                p.append(i)
        else:
            p.append(i)
    f=0
    for i in p:
        del(comp[i-f])
        f += 1
    return mass,comp
def EvalPercent(inp):
    return float(inp.strip('%'))/100.0
def MakeRocks(n):
    rlist = []
    for i in range(1,int(n)+1) if not SingleRock else [1]:
        name = 'Rock '+str(i)
        row1 = [sg.Text('Mass: '),sg.Text('',key=('rock'+str(i)+'mass'),size=6),sg.Text('Invalid',key='rock'+str(i)+'valid',justification='r')]
        header = ['Element','Percent','Volume']
        rows = [['-','-','-'],['-','-','-'],['-','-','-'],['-','-','-'],['-','-','-']]
        table = sg.Table(values=rows,headings=header,key=('rock'+str(i)+'table'),hide_vertical_scroll=True,auto_size_columns=False,col_widths=[6,6,6],size=(0,5),justification='c')
        row3 = sg.Frame('Notes',[[sg.Input(size=20,key=('rock'+str(i)+'notes'))]])
        row4 = [sg.Button('Reset '+str(i)),sg.Checkbox('Fracturing',key='rock'+str(i)+'check')]
        row5 = sg.Image(key='rock'+str(i)+'image')
        rlist.append(sg.Frame(name,[row1,[table],[row3],row4,[row5]],size=(175,330)))
    return rlist
def MakePrimary(u,r,l,s):
    frame = sg.Frame('Primary Info',[
        [sg.Text('Username:'),sg.Input(u,size=31,key='Username')],
        [sg.Text('Rocks Reported:'),sg.Text('{:0>3}'.format(r)),sg.Push(),
         sg.Radio('Planet','loc',key='Planet',default=True if l else False,enable_events=True),
         sg.Radio('Space','loc',key='Space',default=False if l else True,enable_events=True)],
        [sg.Radio('Mining','mode',key='mmode',default=False if s else True,enable_events=True),
         sg.Radio('Scanning','mode',key='smode',default=True if s else False,enable_events=True),
         sg.Push(),sg.Checkbox('Single Rock',key='srm',default=SingleRock)]
        ])
    return frame
def MakeCluster(s,l,d,b,r,v,sh,g,n,a):
    if g:
        frame = sg.Frame('Cluster Info',[
            [sg.Text('Location:'),sg.Input(d,size=5,key='dis'),sg.Text('km'),sg.Input(b,size=3,key='deg'),sg.Text('deg to'),sg.Combo([item for sublist in list(locations.values()) for item in sublist],size=31,key='loc',default_value=l)],
            [sg.Text('Server:'),sg.Combo(regions,size=6,key='reg',default_value=r),sg.Input(v,size=7,key='ver'),sg.Input(sh,size=3,key='shd'),sg.Text('Area:'),sg.Combo(['Flat','Rocky','Mountain','Forest'],size=8,key='area',default_value=a),sg.Push(),sg.Text('Signature:'),sg.Input(s,size=5,key='sig')],
            [sg.Text('Notes:'),sg.Input(n,expand_x=True,key='notes'),sg.Button('Apply')]
            ])
    else:
        frame = sg.Frame('Cluster Info',[
            [sg.Text('Location:'),sg.Combo(spaceloc,key='loc',default_value=l)],
            [sg.Text('Server:'),sg.Combo(regions,size=6,key='reg',default_value=r),sg.Input(v,size=7,key='ver'),sg.Input(sh,size=3,key='shd'),sg.Text('Signature:'),sg.Input(s,size=5,key='sig')],
            [sg.Text('Notes:'),sg.Input(n,size=32,key='notes'),sg.Button('Apply')]
            ])
    return frame
def UpdateRocks():
    global rocks
    for i in [0] if SingleRock else range(int(NumRocks)):
        if i < len(rocks):
            window['rock'+str(i+1)+'mass'].update(str(rocks[i][0]))
            window['rock'+str(i+1)+'valid'].update('Valid')
            newtab = [list((ElementNames[a][0],b,c)) for a,b,c in rocks[i][1]] + [['-','-','-']] * (5-len(rocks[i][1]))
            window['rock'+str(i+1)+'table'].update(newtab)
            window['rock'+str(i+1)+'notes'].update(rocks[i][2])
            window['rock'+str(i+1)+'check'].update(rocks[i][3])
            bio = io.BytesIO()
            rocks[i][4].save(bio,format='PNG')
            window['rock'+str(i+1)+'image'].update(data=bio.getvalue(),visible=True)
        else:
            window['rock'+str(i+1)+'mass'].update('')
            window['rock'+str(i+1)+'valid'].update('Invalid')
            window['rock'+str(i+1)+'table'].update([['-','-','-']]*5)
            window['rock'+str(i+1)+'notes'].update('')
            window['rock'+str(i+1)+'check'].update(False)
            window['rock'+str(i+1)+'image'].update(visible=False)
def BackupGlobal():
    global Username,Type,NumRocks,ClusterNote,Signature,Server,Version,Shard,Location,Distance,Bearing,Area,rocks
    Username = values['Username']
    Type = values['type']
    NumRocks = values['nrocks']
    ClusterNote = values['notes']
    Signature = values['sig']
    Server = process.extractOne(values['reg'],regions)[0]
    Version = values['ver']
    Shard = values['shd']
    Location = ('' if not values['loc'] else process.extractOne(values['loc'],[item for sublist in list(locations.values()) for item in sublist])[0] if Grounded else process.extractOne(values['loc'],spaceloc)[0]) if event != 'Planet' and event !='Space' else ''
    Distance = values['dis'] if event != 'Planet' and event !='Space' else ''
    Bearing = values['deg'] if event != 'Planet' and event !='Space' else ''
    Area = ('' if not values['area'] else process.extractOne(values['area'],['Flat','Rocky','Mountain','Forest'])[0]) if event != 'Planet' and event !='Space' else ''
    rocks = []
def EndSubmit():
    global Username,Type,NumRocks,ClusterNote,Signature,Server,Version,Shard,Location,Distance,Bearing,Area,Applied,rocks
    Username = values['Username']
    Type = ''
    NumRocks = '5'
    ClusterNote = ''
    Signature = ''
    Server = values['reg']
    Version = values['ver']
    Shard = values['shd']
    Location = ''
    Distance = ''
    Bearing = ''
    Area = ''
    Applied = False
    rocks = []

RocksReported = 0
Username = ''
Grounded = True
Keybind = 'alt + f5'
ScanMode = True

NumRocks = '5'
Signature = ''
Type = ''
Location = ''
Distance = ''
Bearing = ''
Server = ''
Version = ''
Shard = ''
ClusterNote = ''
Area = ''
Applied = False
SingleRock = False
Feedback = ''

rocks = [] # structure: (mass,[comp (element,percent,volume)],notes, fracturing, image)

while True:

    RockFrame = sg.Frame('Rock Info',[MakeRocks(NumRocks)])
    PrimaryFrame = MakePrimary(Username,RocksReported,Grounded,ScanMode)
    ClusterFrame = MakeCluster(Signature,Location,Distance,Bearing,Server,Version,Shard,Grounded,ClusterNote,Area)
    CalcFrame = sg.Frame('Signature Calc',[
        [sg.Text('Type:'),sg.Combo(list(Signatures.keys()),size=9,key='type',default_value=Type)],
        [sg.Text('Num. Rocks:'),sg.Push(),sg.Input(NumRocks,size=2,key='nrocks')],
        [sg.Column([[sg.Button('Calculate')]],justification='c')]
        ])
    TutorialsFrame = sg.Frame('How to Use',[[sg.Text('- Fill out all primary information, update keybind if needed')],
                                            [sg.Text('- Fill out all the cluster data; press Apply after')],
                                            [sg.Text('- Press the keybind while scanning each rock to fill in information')]])
    TipsFrame = sg.Frame('Tips',[[sg.Text('- Press ESC to apply keybind changes while in binding mode')],
                                 [sg.Text('- Rock data will be deleted if Apply is pressed again')],
                                 [sg.Text('- Rock data extractor works best at night and on smooth backgrounds')]])
    KeybindFrame = sg.Frame('Keybind',[[sg.Text('Currently:')],
                                       [sg.Text(Keybind,key='keybind')],
                                       [sg.Button('Bind')]],element_justification='c')
    layout = [
    [sg.Text('Welcome to the USG data entry app!')],
    [PrimaryFrame, ClusterFrame, CalcFrame],
    [RockFrame],
    [TutorialsFrame,KeybindFrame,TipsFrame],
    [sg.Text('',key='feedback')],
    [sg.Button('Submit')]]

    window = sg.Window("USG Data App",layout,element_justification='c')
    
    kb.add_hotkey(Keybind, partial(window.write_event_value,'SHOT','SHOT'))
    
    while True:
        event, values = window.read()

        for i in range(len(rocks)):
            rocks[i][2] = values['rock'+str(i+1)+'notes']
            rocks[i][3] = values['rock'+str(i+1)+'check']
        SingleRock = values['srm']

        if event == sg.WIN_CLOSED: break


        elif event == 'Planet' or event == 'Space':
            if Grounded and event == 'Planet': continue
            elif not Grounded and event == 'Space': continue 
            else:
                Grounded = not Grounded
                break


        elif event == 'mmode' or event == 'vmode':
            ScanMode = True if event == 'vmode' else False


        elif event == 'Calculate':
            NumRocks = str(max(min(int(values['nrocks']),10),1))
            Signature = str(int(NumRocks)*Signatures[values['type']])
            Type = values['type']
            window['sig'].update(Signature)


        elif event == 'Apply':
            if not (values['area'] and values['loc'] and values['dis'] and values['deg'] and values['reg'] and values['ver'] and values['shd'] and values['sig']):
                window['feedback'].update("Couldn't apply: Cluster data not filled out")
                continue
            succ = False
            for i,j in Signatures.items():
                if int(values['sig']) % j == 0:
                    values['type'] = i
                    values['nrocks'] = int(values['sig']) // j
                    succ = True
            if not succ:
                window['feedback'].update("Couldn't apply: Signature not correct")
                continue
            BackupGlobal()
            Applied = True
            break


        elif event == 'Bind':
            rec = kb.record(until='Esc')
            temp = []
            for i in range(len(rec)-1):
                rec[i] = str(rec[i])[14:str(rec[i]).index(' ')]
                if rec[i] not in temp:
                    temp.append(rec[i])
            Keybind = ' + '.join(temp)
            window['keybind'].update(Keybind)
            kb.unhook_all_hotkeys()
            kb.add_hotkey(Keybind, partial(window.write_event_value,'SHOT','SHOT'))


        elif event == 'SHOT':
            if len(rocks) == int(NumRocks) or (len(rocks) == 1 and SingleRock):
                window['feedback'].update("Couldn't get rock: Rocks data full")
                continue
            if not (values['area'] and values['loc'] and values['dis'] and values['deg'] and values['reg'] and values['ver'] and values['shd'] and values['sig'] and Applied):
                window['feedback'].update("Couldn't get rock: Cluster data not applied")
                continue
            try:
                image = ScreenCap('Star Citizen')
            except Exception as e:
                window['feedback'].update('Star Citizen Not Found')
                continue
            try:
                mass,comp = VMImage(np.array(image)) if ScanMode else MMImage(np.array(image))
            except Exception as e:
                window['feedback'].update('Error during text extraction: '+str(e))
                continue
            try:
                if abs(sum([EvalPercent(y) for x,y in comp])-1) > .01: raise Exception('Percents not 1')
                volume = mass/sum([EvalPercent(y)*ElementNames[x][1] for x,y in comp])
                comp = [(x,y,'{:.2f}'.format(EvalPercent(y)*volume)) for x,y in comp]
                image = image.resize((int(image.size[0]/10),int(image.size[1]/10)))
                rocks.append([mass,comp,'',False,image])
            except Exception as e:
                window['feedback'].update("Couldn't interpret composition: "+str(e))
                continue
            try:
                UpdateRocks()
                window['feedback'].update('Rock input successfully!')
            except Exception as e:
                window['feedback'].update("Couldn't show data: "+str(e))
                continue


        elif event[:5] == 'Reset':
            active = int(event[-1])-1
            if active < len(rocks):
                try:
                    del(rocks[active])
                    UpdateRocks()
                except Exception as e:
                    window['feedback'].update("Couldn't delete rock: "+str(e))


        elif event == 'Submit':
            if len(rocks) == int(NumRocks) or (len(rocks) == 1 and SingleRock):

                planet = [x for x,y in locations.items() if values['loc'] in y][0] if Grounded else values['loc']

                try:
                    for i in rocks:
                        data = {
                            "entry.107658724": values['Username'],
                            "entry.1005637855": values['reg'],
                            "entry.2137031800": values['ver'],
                            "entry.1970416187": values['shd'],
                            "entry.816697130": planet,
                            "entry.185553886": values['sig'],
                            "entry.1049853798": str(i[0]), # mass
                            "entry.577072814": values['type'],
                            "entry.923538061": values['loc'],
                            "entry.998639438": values['dis'] if Grounded else '0',
                            "entry.797873939": values['deg'] if Grounded else '0',
                            "entry.384703431": values['area'] if Grounded else 'Flat',
                            "entry.144604195": values['notes'] + '///' + i[2] if values['notes'] and i[2] else i[2] if i[2] else values['notes'], # notes
                            "draftResponse": [],
                            "pageHistory": 0
                            }
                        for j in i[1][:-1]:
                            data['entry.'+ElementNames[j[0]][2]] = j[1] # element detection
                        if i[3]: data["entry.1506299456"] = 'TRUE' # fracturing detection
                        print(data)
                        print(requests.post('https://docs.google.com/forms/d/e/1FAIpQLSdiVhyu3uHbpOvq8vaV7SA59nUh8jDONe-zQjW1qPyMf_81zg/formResponse',data=data))
                        sleep(.2)
                    window['feedback'].update("Rock submitted successfully!" if SingleRock else "Cluster submitted successfully!")
                    EndSubmit()
                    break
                except Exception as e:
                    window['feedback'].update("Submit Failed: "+str(e))
            else:
                window['feedback'].update("Couldn't submit: Rocks data not full")
        print(event)
        print(values)


    kb.unhook_all_hotkeys()
    # data would save here
    window.close()
    if event == sg.WIN_CLOSED: break