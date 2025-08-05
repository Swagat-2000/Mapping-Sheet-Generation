import warnings
import pandas as pd
import decimal
#import random
#import xlsxwriter
from ArxmlParser.NoxArxmlESISParser import get_arxml_ESIS_data
from ArxmlParser.NoxArxmlESISParser import updatetobalance

arxmlFile=r'SmartDnc_EcuExtract_JLR_DCU_MY28.arxml'
NoxNtwDoc='NetworkDocumentation_CAN_PCM_DCU_Cluster_TAchange.xlsx' 
NoxMapping= 'JLR_MY28_AJ20D6_PCM_DCU_CAN_MappingSheet_v1.xlsx'
esisSpec=None
projectselection='DCU' # NOX or DCU
selectopforrow = 'testteam'  # testteam / ntwdoc

outFile='PCM_DCU_21july2025_V3.xlsx' #//PCM_Nox_NC11_20may25_V3
ArxmlExpo =outFile
outputfile='TestMappingSheet_NoxCAN_NC11_V2.xlsx'

warnings.filterwarnings('ignore')  # Suppress all warnings
get_arxml_ESIS_data(arxmlFile,esisSpec,outFile)
updatetobalance(outFile)

warnings.filterwarnings('ignore')  # Suppress all warnings
get_arxml_ESIS_data(arxmlFile,esisSpec,outFile)
updatetobalance(outFile)
    
xls = pd.ExcelFile(NoxNtwDoc)
df_NoxNtwDoc = pd.read_excel(xls, sheet_name=xls.sheet_names[0]) 
df_NoxMapping=pd.read_excel(NoxMapping) 
df_ArxmlExpo=pd.read_excel(ArxmlExpo) 
ND_Direction  = df_NoxNtwDoc['Direction'].tolist()
ND_Frame  = df_NoxNtwDoc['Frame'].tolist()
ND_PDU  = df_NoxNtwDoc['PDU'].tolist()
ND_Signal  = df_NoxNtwDoc['Signal'].tolist()
ND_Length  = df_NoxNtwDoc['Length'].tolist()
ND_UpdateBit  = df_NoxNtwDoc['Update Bit'].tolist()
ND_ASW  = df_NoxNtwDoc['ASW'].tolist()
ND_NetworkMin  = df_NoxNtwDoc['Network Min'].tolist()
ND_NetworkMax  = df_NoxNtwDoc['Network Max'].tolist()
ND_Resolution  = df_NoxNtwDoc['Resolution'].tolist()
ND_Offset  = df_NoxNtwDoc['Offset'].tolist()
ND_PhysicalMin  = df_NoxNtwDoc['Physical Min'].tolist()
ND_PhysicalMax  = df_NoxNtwDoc['Physical Max'].tolist()
ND_IniCalibration  = df_NoxNtwDoc['Ini Calibration'].tolist()
ND_DefaultCalibration  = df_NoxNtwDoc['Default Calibration'].tolist()
try:
    ND_PDUTimeoutDSQ  = df_NoxNtwDoc['PDU Timeout DSQ'].tolist()
except:
    ND_PDUTimeoutDSQ  = df_NoxNtwDoc['PDU Timeout DFC'].tolist()
ND_PDUTimeoutDebounceCalibration  = df_NoxNtwDoc['PDU Timeout Debounce Calibration'].tolist()

# Needed this to add ND_PDU_Modified to match with DEV_PDU
ND_PDU_Modified = [item if '_pdu' in item else item + '_pdu' for item in ND_PDU]

print(ND_PDU_Modified)

Ar_PDU  = df_ArxmlExpo['PDU'].tolist()
Ar_Frame  = df_ArxmlExpo['Frame'].tolist()
Ar_Signal1  = df_ArxmlExpo['Signal'].tolist() # isignal
Ar_CompuMethod  = df_ArxmlExpo['Compu Method'].tolist()
Ar_SystemSignal  = df_ArxmlExpo['System Signal'].tolist()
Ar_InitValue  = df_ArxmlExpo['Init Value'].tolist()
Ar_Range  = df_ArxmlExpo['Range1'].tolist()
Ar_Offset  = df_ArxmlExpo['Offset1'].tolist()
Ar_Factor  = df_ArxmlExpo['Factor1'].tolist()
Ar_FrameID  = df_ArxmlExpo['Frame ID'].tolist()
Ar_SignalLength  = df_ArxmlExpo['Signal Length'].tolist()
Ar_Sender  = df_ArxmlExpo['Sender'].tolist() # ECM --> PCM_P2

import re

if projectselection == 'DCU':
    def clean_ar_signal_name(signal):
        return re.sub(r'^_+([^_]*_)?', '', signal)
    Ar_Signal = [clean_ar_signal_name(s) for s in Ar_Signal1]
    # Normalize AR_Signal by stripping known prefixes
    def strip_prefix(s):
        prefixes = ['NW_COM_', 'NW_Com_', 'NW_Tx_', 'NW_SCRT_', 'NW_', 'Tx_', '']
        for p in prefixes:
            if s.startswith(p):
                return s[len(p):]
        return s
    converted_AR = []
    for sig in Ar_Signal:
        stripped = strip_prefix(sig)
        match = next((nd for nd in ND_Signal if nd.endswith(stripped)), None)
        if match:
            converted_AR.append(match)
    print("Matched AR_Signal in ND_Signal format:")
    for sig in converted_AR:
        print(sig)
    Ar_Signal = converted_AR
else:
    Ar_Signal = Ar_Signal1

DevDoc_FrameName  = df_NoxMapping['FrameName'].tolist()
try:
    DevDoc_PDU  = df_NoxMapping['PDU'].tolist()
except:
    DevDoc_PDU  = df_NoxMapping['PduName'].tolist()
DevDoc_NetworkSignalName  = df_NoxMapping['NetworkSignalName'].tolist()


DevDoc_ApplicationSignalName  = df_NoxMapping['ApplicationSignalName'].tolist()
DevDoc_Cycle  = df_NoxMapping['Cycle'].tolist()
DevDoc_QFSignalName  = df_NoxMapping['QF Signal Name'].tolist()
DevDoc_Invalid_Range  = df_NoxMapping['Invalid_Range'].tolist()
try:
    DevDoc_DSQforInvalidrange  = df_NoxMapping['DSQ for Invalid range'].tolist()
except:
    DevDoc_DFCforInvalidrange  = df_NoxMapping['DFC for Invalid range'].tolist()
    
def normalize_network_signal_name(name):
    return (name.replace('NW_Tx_Com_', 'Com_')
                .replace('NW_Com_', 'Com_')
                .replace('NW_Tx_', '')
                .replace('NW_', ''))

# Only normalize DevDoc_NetworkSignalName if it doesn't match ND_Signal
if not any(sig in ND_Signal for sig in DevDoc_NetworkSignalName):
    DevDoc_NetworkSignalName = [normalize_network_signal_name(name) for name in DevDoc_NetworkSignalName]

MDmergerDevDocASW = [''] * len(ND_Signal)
MDmergerDevDocCyclic = [''] * len(ND_Signal)
TT = 0
check = 0
for loopvar in range(len(ND_Signal)):
    check = 1
    CC = 1
    for loopvar1 in range(len(DevDoc_NetworkSignalName)):
        if check == 1 or (CC == 1) or (CC == 0):
            if DevDoc_NetworkSignalName[loopvar1].lower() == ND_Signal[loopvar].lower():
                if DevDoc_PDU[loopvar1].lower() == ND_PDU_Modified[loopvar].lower():
                    if DevDoc_FrameName[loopvar1].lower() == ND_Frame[loopvar].lower():
                        check = 0
                        if DevDoc_ApplicationSignalName[loopvar1].lower() == ND_ASW[loopvar].lower():
                            CC = 2
                            TT = TT + 1
                            MDmergerDevDocASW[loopvar] = DevDoc_ApplicationSignalName[loopvar1]
                            MDmergerDevDocCyclic[loopvar] = DevDoc_Cycle[loopvar1]
                        else:
                            if CC == 2:
                                CC = 2
                            else:
                                CC = 0
    if check == 1:
        print('interface not matching')
    if CC == 0:
        print('not mactching asw')
print('Total no of interface in development input: ', TT)

loopvar=0
loopvar1=0                    
NtMergAr_PDU  =['']*len(ND_Signal)
NtMergAr_Frame  =['']*len(ND_Signal)
NtMergAr_Signal  =['']*len(ND_Signal)
NtMergAr_CompuMethod  =['']*len(ND_Signal)
NtMergAr_SystemSignal  =['']*len(ND_Signal)
NtMergAr_InitValue  =['']*len(ND_Signal)
NtMergAr_Range  =['']*len(ND_Signal)
NtMergAr_Offset  =['']*len(ND_Signal)
NtMergAr_Factor  =['']*len(ND_Signal)
NtMergAr_FrameID =['']*len(ND_Signal)
NtMergAr_SignalLength  =['']*len(ND_Signal)
NtMergAr_Sender  =['']*len(ND_Signal) 
SigRangeMin=['']*len(ND_Signal) 
SigRangeMax=['']*len(ND_Signal) 
TT=0
for loopvar in range(len(ND_Signal)):
    for loopvar1 in range(len(Ar_Signal)):
        if Ar_Signal[loopvar1].lower()==ND_Signal[loopvar].lower():
            if Ar_PDU[loopvar1].lower()==ND_PDU[loopvar].lower():
                if Ar_Frame[loopvar1].lower()==ND_Frame[loopvar].lower():
                    TT=TT+1  
                    #print("match between:" ,Ar_Signal[loopvar1],ND_Signal[loopvar])
                    NtMergAr_PDU[loopvar]  =Ar_PDU[loopvar1]
                    NtMergAr_Frame[loopvar]  =Ar_Frame[loopvar1]
                    NtMergAr_Signal[loopvar]  =Ar_Signal[loopvar1] 
                    print(f"NtMergAr_Signal[{loopvar}] = {NtMergAr_Signal[loopvar]}, Ar_Signal[{loopvar1}] = {Ar_Signal[loopvar1]}")
                    NtMergAr_CompuMethod[loopvar]  =Ar_CompuMethod[loopvar1]
                    NtMergAr_SystemSignal[loopvar]  =Ar_SystemSignal[loopvar1]
                    NtMergAr_InitValue[loopvar]  =Ar_InitValue[loopvar1]
                    NtMergAr_Range[loopvar]  =Ar_Range[loopvar1]
                    tt=Ar_Range[loopvar1].split('-')
                    SigRangeMin [loopvar]= tt[0]
                    SigRangeMax [loopvar]= tt[-1]
                    NtMergAr_Offset[loopvar]  =Ar_Offset[loopvar1]
                    NtMergAr_Factor[loopvar]  =Ar_Factor[loopvar1] 
                    NtMergAr_FrameID[loopvar] =Ar_FrameID[loopvar1]
                    NtMergAr_SignalLength[loopvar]  =Ar_SignalLength[loopvar1]
                    #if Ar_Sender [loopvar1] !='ECM' or Ar_Sender [loopvar1] !='ecb' or Ar_Sender [loopvar1] == 'Nan' or Ar_Sender [loopvar1] =='' or Ar_Sender [loopvar1] =='nan':
                    if ND_Direction[loopvar].lower()=='tx':
                        NtMergAr_Sender [loopvar] ='PCM_P2Nox'
                    else:
                        NtMergAr_Sender [loopvar] = 'Vector_XXX'
        #else:
         #   print("mismtach between:" ,Ar_Signal[loopvar1],ND_Signal[loopvar])
print('Total no of interface in mergerd with development input from arxml: ',TT)

    
seen_PduName = set()
uniqPduName = []
TA_PduNumber = []
count=2000
for forloop_uniqPduName in NtMergAr_PDU:     
    if forloop_uniqPduName not in seen_PduName:
        uniqPduName.append((forloop_uniqPduName))           
        seen_PduName.add(forloop_uniqPduName) 
        count=count+1  
    TA_PduNumber.append(count)
#print(TA_PduNumber)

TT = 0
for loopvar in range(len(ND_Signal)):
    found = False
    for loopvar1 in range(len(Ar_Signal)):
        if Ar_Signal[loopvar1].lower() == ND_Signal[loopvar].lower():
            NtMergAr_PDU[loopvar] = Ar_PDU[loopvar1]
            NtMergAr_Frame[loopvar] = Ar_Frame[loopvar1]
            NtMergAr_Signal[loopvar] = Ar_Signal[loopvar1]
            NtMergAr_CompuMethod[loopvar] = Ar_CompuMethod[loopvar1]
            NtMergAr_SystemSignal[loopvar] = Ar_SystemSignal[loopvar1]
            NtMergAr_InitValue[loopvar] = Ar_InitValue[loopvar1]
            NtMergAr_Range[loopvar] = Ar_Range[loopvar1]
            tt = Ar_Range[loopvar1].split('-')
            SigRangeMin[loopvar] = tt[0]
            SigRangeMax[loopvar] = tt[-1]
            NtMergAr_Offset[loopvar] = Ar_Offset[loopvar1]
            NtMergAr_Factor[loopvar] = Ar_Factor[loopvar1]
            NtMergAr_FrameID[loopvar] = Ar_FrameID[loopvar1]
            NtMergAr_SignalLength[loopvar] = Ar_SignalLength[loopvar1]
            if ND_Direction[loopvar].lower() == 'tx':
                NtMergAr_Sender[loopvar] = 'PCM_P2Nox'
            else:
                NtMergAr_Sender[loopvar] = 'Vector_XXX'
            TT += 1
            found = True
            print(f"NtMergAr_Signal[{loopvar}] = {NtMergAr_Signal[loopvar]}, Ar_Signal[{loopvar1}] = {Ar_Signal[loopvar1]}")
            break  # Stop after first match
    if not found:
        # Optionally fill with defaults
        NtMergAr_PDU[loopvar] = ''
        NtMergAr_Frame[loopvar] = ''
        NtMergAr_Signal[loopvar] = ND_Signal[loopvar]
        NtMergAr_CompuMethod[loopvar] = ''
        NtMergAr_SystemSignal[loopvar] = ''
        NtMergAr_InitValue[loopvar] = ''
        NtMergAr_Range[loopvar] = ''
        SigRangeMin[loopvar] = ''
        SigRangeMax[loopvar] = ''
        NtMergAr_Offset[loopvar] = ''
        NtMergAr_Factor[loopvar] = ''
        NtMergAr_FrameID[loopvar] = ''
        NtMergAr_SignalLength[loopvar] = ''
        NtMergAr_Sender[loopvar] = ''
print('Total no of interface in mergerd with development input from arxml:', TT)

sseen_PduName = set()
uniqPduName = []
TA_PduNumber = []
count=2000
for forloop_uniqPduName in NtMergAr_PDU:     
    if forloop_uniqPduName not in seen_PduName:
        uniqPduName.append((forloop_uniqPduName))           
        seen_PduName.add(forloop_uniqPduName) 
        count=count+1  
    TA_PduNumber.append(count)
#print(TA_PduNumber)


busResolution=[]
BusOffset=[]
BusMin=[]
BusMax=[]
TA_SigLength=[]

if selectopforrow.lower()=='testteam':
    busResolution=   NtMergAr_Factor #SigResolution
    BusOffset = NtMergAr_Offset #SigOffset
    BusMin = SigRangeMin #'SigPhyMin'#
    BusMax = SigRangeMax #'SigPhyMax'#
    TA_SigLength =NtMergAr_SignalLength #'SigLeninbit'
else:
    busResolution=   ND_Resolution #TA_Resolution
    BusOffset = ND_Offset #TA_Offset
    BusMin = ND_NetworkMin #TA_NtwMin
    BusMax = ND_NetworkMax #TA_NtwMax
    TA_SigLength = ND_Length #TA_SigLength
    
TA_PhySig_ValRngVal=[]
TA_RawSig_ValRngVal=[]
# TA_PhySig_ValRngVal=['']*len(ND_Signal) 


# TA_RawSig_ValRngVal=['']*len(ND_Signal) 

validRange_samples_req = 15
for loopvar3 in range(len(NtMergAr_Factor)):
    TT=0
    #print(type(len))
    #print("line no 226")
    Valid_raw_min=BusMin[loopvar3]
    Valid_raw_max=BusMax[loopvar3]    
    valid_raw_values=[]
    #print(Valid_raw_min,Valid_raw_max)
    if Valid_raw_max == '' or Valid_raw_max is None:
        Valid_raw_max = 0
    valid_raw_values = range(int(Valid_raw_max)+1)
    if busResolution[loopvar3] == '' or busResolution[loopvar3] is None:
        factor = 0.0
    else:
        factor = float(busResolution[loopvar3])
    if BusOffset[loopvar3] == '' or BusOffset[loopvar3] is None:
        offset = 0.0
    else:
         offset=float(BusOffset[loopvar3])
    d=decimal.Decimal(str(factor))
    dec_places=-(d.as_tuple().exponent)
    try:
        Valid_raw_max=int(Valid_raw_max)
    except:
        Valid_raw_max=Valid_raw_max
    if Valid_raw_max<=1073741828:
        sig_length_str = TA_SigLength[loopvar3]
        if sig_length_str == '' or sig_length_str is None:
            sig_length = 0
        else:
            sig_length = int(sig_length_str)
        maxval = pow(2, sig_length)        
    else:
        maxval= valid_raw_values[-1]
        
    if maxval == (Valid_raw_max+1):
        Valid_raw_list1=[]
        Valid_phy_list1=[]
        samples_ValRng=(valid_raw_values[-1]-valid_raw_values[0])
        #print(samples_ValRng)
        if(samples_ValRng > validRange_samples_req): # out of 15 samples pick 3 samples at start,mid,end
            sample_partition=validRange_samples_req//3
            for i in range(sample_partition):
                Valid_raw_list1.append((int(valid_raw_values[i])))
                Valid_raw_list1.append((int(valid_raw_values[-sample_partition+i])))
                Valid_raw_list1.append((int((valid_raw_values[i]+valid_raw_values[-sample_partition+i])//2)))
            Valid_raw_list1.sort()
            #print(len(Valid_raw_list1))
            for i in range(len(Valid_raw_list1)):
                Valid_phy_list1.append(round((Valid_raw_list1[i]*factor)+offset,dec_places))
            #print(Valid_raw_list1)
            #print(Valid_phy_list1)
        else:
            Valid_raw_list1=valid_raw_values
            phy_val_lst=[]
            for i in range(len(Valid_raw_list1)):
                Valid_phy_list1.append(round((Valid_raw_list1[i]*factor)+offset,dec_places))
            #print(len(Valid_raw_list1))
           # print(Valid_phy_list1)
            #print(Valid_raw_list1)
       # tempraw=[]
       # tempphy=[]
       # tempraw=(*Valid_phy_list1, sep=', ') 
        TA_PhySig_ValRngVal.append(Valid_phy_list1)
        TA_RawSig_ValRngVal.append(list(Valid_raw_list1))   
        #print(TA_RawSig_ValRngVal[loopvar3])
    else:
        print("the length is greather than the max value")
loopvar=0
for loopvar in range(len(TA_PhySig_ValRngVal)):   
   # print(type(TA_RawSig_ValRngVal[loopvar]),type(TA_RawSig_ValRngVal[loopvar]))
    try:
        aa=str(TA_RawSig_ValRngVal[loopvar])
        aa1=aa.replace("[","")
        aa2=aa1.replace("]","")
        TA_RawSig_ValRngVal[loopvar] = aa2
        bb=str(TA_PhySig_ValRngVal[loopvar])
        bb1=bb.replace("[","")
        bb2=bb1.replace("]","")
        TA_PhySig_ValRngVal[loopvar] = bb2
    except:
        TA_RawSig_ValRngVal[loopvar] = TA_RawSig_ValRngVal[loopvar] 
        TA_PhySig_ValRngVal[loopvar] =TA_PhySig_ValRngVal[loopvar]
        
import pandas as pd

# Your initialization
SignalGrpOrdNo=['']*len(ND_Signal) 
SignalGrpName=['']*len(ND_Signal) 
SignalGrpSize=['']*len(ND_Signal) 
SigGrpStartByte=['']*len(ND_Signal) 
SigGrpStopByte=['']*len(ND_Signal) 
SigGrpStartByte_mask=['']*len(ND_Signal) 
SigGrpStopByte_mask=['']*len(ND_Signal) 
DSQforInvalid=['']*len(ND_Signal) 
Invalid_Range=['']*len(ND_Signal) 
QFSignalName=['']*len(ND_Signal) 
DSQ=['']*len(ND_Signal) 
CONV_RULE=['']*len(ND_Signal) 
TA_SigMask=['']*len(ND_Signal) 
ASW_Bit_Field=['']*len(ND_Signal) 
OutofRangeValues =['']*len(ND_Signal) 
P_ECUFaultTO_TimeOutlabel=['']*len(ND_Signal) 
P_ECUFaultCHK_Checksumlabel=['']*len(ND_Signal) 
P_ECUFaultCTR_Alivecounterlabel=['']*len(ND_Signal) 
PlausOKSignal=['']*len(ND_Signal) 
LastAliveSignal=['']*len(ND_Signal) 
Byte0_Const=['']*len(ND_Signal) 
Byte1_Const=['']*len(ND_Signal) 
NodeMon_DFC=['']*len(ND_Signal) 
Base_Repet_FR=['']*len(ND_Signal) 
NodeMon_DFC_debdef=['']*len(ND_Signal) 
NodeMon_DFC_debok=['']*len(ND_Signal) 
SigCategory=['']*len(ND_Signal) 
GateWayPDU=['']*len(ND_Signal) 
GateWaySignal=['']*len(ND_Signal) 
SigGrpUB=['']*len(ND_Signal) 
SigUB=['']*len(ND_Signal) 
SignalGrpNo=['']*len(ND_Signal) 
TAGID=['']*len(ND_Signal) 
TA_PlausoffCAl=['']*len(ND_Signal) 
TA_PlausOK=['']*len(ND_Signal) 
TA_CHKDFC=['']*len(ND_Signal) 
TA_CTRDFC=['']*len(ND_Signal) 
TA_PlausFID=['']*len(ND_Signal) 
TA_CHKDebDef=['']*len(ND_Signal) 
TA_CHKDebOk=['']*len(ND_Signal) 
TA_CTRDebDef=['']*len(ND_Signal) 
TA_CTRDebOk=['']*len(ND_Signal) 
TA_CTRStuckMaxCalib=['']*len(ND_Signal) 
TA_RawMessage=['']*len(ND_Signal) 
BusUnit=['']*len(ND_Signal) 
TA_SSigBitField=['']*len(ND_Signal) 
TA_QFDFC=['']*len(ND_Signal) 
TA_InvalidRangDSQ=['']*len(ND_Signal) 
TA_QFDFCDebDef=['']*len(ND_Signal) 
TA_QFDFCDebOk=['']*len(ND_Signal) 
TA_CHKCalc=['']*len(ND_Signal) 
TA_CTRDiff=['']*len(ND_Signal) 
TA_CTRLstVal=['']*len(ND_Signal) 
TA_SigUpdateBit=['']*len(ND_Signal) 
TA_PhySig_InValRngVal=['']*len(ND_Signal) 
TA_RawSig_InValRngVal=['']*len(ND_Signal) 
TA_PhySig_OutofRngVal=['']*len(ND_Signal) 
TA_RawSig_OutofRngVal=['']*len(ND_Signal) 

# All columns in dictionary
data_dict = {
    "FrameName": ND_Frame,
    "PduName": ND_PDU,
    "NetworkSignalName": ND_Signal,
    "Compumethod(Application)": NtMergAr_CompuMethod,
    "ApplicationSignalName": ND_ASW,
    "Ntw_Direction": NtMergAr_Sender,
    "Cycle": MDmergerDevDocCyclic,
    "isignal": NtMergAr_Signal,
    "SigIntVal": NtMergAr_InitValue,
    "SigValRange": NtMergAr_Range,
    "SigRangeMin": SigRangeMin,
    "SigRangeMax": SigRangeMax,
    "SigOffset": NtMergAr_Offset,
    "SigResolution": NtMergAr_Factor,
    "SigLeninbit": NtMergAr_SignalLength,
    "TA_NtwMin": ND_NetworkMin,
    "TA_NtwMax": ND_NetworkMax,
    "TA_Resolution": ND_Resolution,
    "TA_Offset": ND_Offset,
    "TA_PhyMin": ND_PhysicalMin,
    "TA_PhyMax": ND_PhysicalMax,
    "TA_IniCal": ND_IniCalibration,
    "TA_DefCal": ND_DefaultCalibration,
    "TA_DSQ_TO": ND_PDUTimeoutDSQ,
    "TA_DSQ_TO_DEB": ND_PDUTimeoutDebounceCalibration,
    "TA_SigLength": ND_Length,
    "TA_Direction": ND_Direction,
    "SignalGrpOrdNo": SignalGrpOrdNo,
    "SignalGrpName": SignalGrpName,
    "SignalGrpSize": SignalGrpSize,
    "SigGrpStartByte": SigGrpStartByte,
    "SigGrpStopByte": SigGrpStopByte,
    "SigGrpStartByte_mask": SigGrpStartByte_mask,
    "SigGrpStopByte_mask": SigGrpStopByte_mask,
    "DSQ for Invalid range": DSQforInvalid,
    "Invalid_Range": Invalid_Range,
    "QFSignalName": QFSignalName,
    "DSQ": DSQ,
    "CONV_RULE": CONV_RULE,
    "TA_SigMask": TA_SigMask,
    "ASW_Bit_Field": ASW_Bit_Field,
    "Out of Range Values ": OutofRangeValues,
    "P_ECUFaultTO - TimeOut label": P_ECUFaultTO_TimeOutlabel,
    "P_ECUFaultCHK - Checksum label": P_ECUFaultCHK_Checksumlabel,
    "P_ECUFaultCTR - Alive counter label": P_ECUFaultCTR_Alivecounterlabel,
    "PlausOKSignal": PlausOKSignal,
    "LastAliveSignal": LastAliveSignal,
    "Byte0_Const": Byte0_Const,
    "Byte1_Const": Byte1_Const,
    "NodeMon_DFC": NodeMon_DFC,
    "Base_Repet_FR": Base_Repet_FR,
    "NodeMon_DFC_debdef": NodeMon_DFC_debdef,
    "NodeMon_DFC_debok": NodeMon_DFC_debok,
    "SigCategory": SigCategory,
    "GateWayPDU": GateWayPDU,
    "GateWaySignal": GateWaySignal,
    "SigGrpUB": SigGrpUB,
    "SigUB": SigUB,
    "SignalGrpNo": SignalGrpNo,
    "TAGID": TAGID,
    "TA_PlausoffCAl": TA_PlausoffCAl,
    "TA_PlausOK": TA_PlausOK,
    "TA_CHKDFC": TA_CHKDFC,
    "TA_CTRDFC": TA_CTRDFC,
    "TA_PlausFID": TA_PlausFID,
    "TA_CHKDebDef": TA_CHKDebDef,
    "TA_CHKDebOk": TA_CHKDebOk,
    "TA_CTRDebDef": TA_CTRDebDef,
    "TA_CTRDebOk": TA_CTRDebOk,
    "TA_CTRStuckMaxCalib": TA_CTRStuckMaxCalib,
    "TA_RawMessage": TA_RawMessage,
    "BusUnit": BusUnit,
    "TA_SSigBitField": TA_SSigBitField,
    "TA_QFDFC": TA_QFDFC,
    "TA_InvalidRangDSQ": TA_InvalidRangDSQ,
    "TA_QFDFCDebDef": TA_QFDFCDebDef,
    "TA_QFDFCDebOk": TA_QFDFCDebOk,
    "TA_CHKCalc": TA_CHKCalc,
    "TA_CTRDiff": TA_CTRDiff,
    "TA_CTRLstVal": TA_CTRLstVal,
    "TA_SigUpdateBit": TA_SigUpdateBit,
    "TA_PhySig_InValRngVal": TA_PhySig_InValRngVal,
    "TA_RawSig_InValRngVal": TA_RawSig_InValRngVal,
    "TA_PhySig_OutofRngVal": TA_PhySig_OutofRngVal,
    "TA_RawSig_OutofRngVal": TA_RawSig_OutofRngVal,
    "TA_PhySig_ValRngVal": TA_PhySig_ValRngVal,
    "TA_RawSig_ValRngVal": TA_RawSig_ValRngVal,
    "TA_PduNumber": TA_PduNumber
}

# Print lengths and highlight mismatches
expected_len = len(ND_Signal)
print(f"Expected length for all columns: {expected_len}")
for key, value in data_dict.items():
    l = len(value)
    if l != expected_len:
        print(f"Length mismatch: {key}: {l} (expected {expected_len})")
    else:
        print(f"{key}: {l}")

# Optionally: Auto-fix mismatches by padding or trimming
for key, value in data_dict.items():
    if len(value) < expected_len:
        print(f"Padding {key} from {len(value)} to {expected_len}")
        data_dict[key] = value + [''] * (expected_len - len(value))
    elif len(value) > expected_len:
        print(f"Trimming {key} from {len(value)} to {expected_len}")
        data_dict[key] = value[:expected_len]

# Now create the DataFrame
df = pd.DataFrame(data_dict)
print("DataFrame shape:", df.shape)
df.to_excel(outputfile, header=True, index=False)
