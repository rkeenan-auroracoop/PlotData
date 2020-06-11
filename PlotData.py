import os
import pprint
# import openpyxl
from openpyxl import load_workbook

# This is for the PLANTING FORM

docDir = r'C:\Users\rkeenan\OneDrive - Aurora Cooperative\Documents\2020 Tableau Updates\Al Perry'

plantingWriteFile = (
    r'C:\Users\rkeenan\OneDrive - Aurora Cooperative\Documents\2020 Tableau Updates\Al Perry\Test Plot(PLANTING).txt', 'w')

PV_113_V89 = ['113-V89', '113-V89 VT2', '113-V89 VT2P', '113-V89 VT2PRIB', '113-V89VT2', '113-V89VT2PRIB', '113V89 VT2', 'PV 113 V89 VT2', 'PV 113-V89',
              'PV 113-V89 - VT2', 'PV 113-V89 VT2PRIB 32k', 'PV113 V89 VT2PRIB', 'PV113-V89', 'PV113-V89 VT2', 'PV113-V89 VT2RIB- Check', 'PV113-V89-VT2PRIB', 'PV113-V89VT2']

NK_1082_5222A = ['1082 5222', '1082-5222', '1082-5222A', '1082A 5222',
                 'NK1082-5222A', 'NK1082-5222A Brand', 'NK1082-5222AEZ1']

PV_110_H20SS = ['110-H20 SS', '110-H2O STX']

PV_110_H20VT2PRIB = ['110-H20 VT2', '110-H20 VT2P', '110-H20 VT2PRIB', '110-H20VT2', '110-H20VT2PRIB', 'PV110 H20 VT2PRIB',
                     'PV110-H20 VT2', 'PV110-H20 VT2PRIB', 'PV110-H20 VT2PRIB 34k', 'PV110-H20 VT2RIB', 'PV110-H20VT2', 'PV110 H2O VT2PRIB', 'PV110-H20-VT2PRIB', 'PV110H20-VT2']
PV_112_Q70BT2PRIB = ['112-Q70', '112-Q70 VT2',
                     '112-Q70 VT2P', '112-Q70 VT2PRIB', '112-Q70VT2', 'PV 112- Q70', 'PV 112-Q70 - VT2', 'PV 112-Q70 VT2', 'PV112 Q70 VT2PRIB', 'PV112-Q70', 'PV112-Q70 VT2', 'PV112-Q70 VT2PRIB', 'PV112-Q70 VT2PRIB 36k', 'PV112-Q70TT2', 'PV112Q-70']

PV_112_T69SS = ['112-T69 SS', '112-T69 SSTX',
                '112-T69 SSTX RIB', '112-T69 STX', '112-T69SSRIB', 'PV 112-T69 - VT2', 'PV112 T69 SSRIB', 'PV112-T69', 'PV112-T69 SS', 'PV112-T69 SSRIB', 'PV112-T69 SSRIB 36k']
PV_113_B40 = ['113-B40 DGVT2', '113-B40 DGVT2PRIB',
              '113-B40 VT2', '113-B40 VT2P', '113-B40DGVT2', 'PV113 B40 DGVT2PRIB', 'PV113-B40', 'PV113-B40 DGVT2PRIB', 'PV113-B40 VT2', 'PV113-B40DGVT2']
PV_115_D59 = ['115-D59 VT2', '115-D59 VT2P', '115-D59 VT2P', '115-D59VT2PRIB', '115D59',
              '115D59 VT2', 'PV 115-D59 - VT2', 'PV 115-D59 VT2', 'PV 115-D59 VT2PRIB 34k', 'PV 115-D59VT2', 'PV115 D59 VT2PRIB', 'PV115-D59', 'PV115-D59 VT2', 'PV115-D59 VT2PRIB', 'PV115-D59 VT2RIB', 'PV115-D59TV2', 'PV115-D59VT2']
PV_115_M60 = ['115-M60 TRE RIB', '115-M60 TRERIB', '115-M60 Tr', '115-M60TRE RIB', '115-M60TRERIB', '115-R60 Tre', '115M60', '115M60 VT2',
              'PV 115-M60 TRECEPTA 34k', 'PV115 M60 TRERIB', 'PV115-M60', 'PV115-M60 TRE', 'PV115-M60 TRERIB', 'PV115-M60 TreceptraRIB', 'PV115-M60VT2', 'PV115M60 TRERIB']
PV_114_R50_VT2PRIB = ['114-R50 VT2', '114-R50 VT2P',
                      '114-R50VT2PRIB', '114R50 VT2', 'PV 114 R50 VT2', 'PV 114-R50 VT2PRIB 32k', 'PV114 R50 VT2PRIB', 'PV114-R50 VT2', 'PV114-R50 VT2PRIB', 'PV114-R50-VT2RIB', 'PV114-R50VT2', 'PV114R50VT2', 'PV114-R50-VT2PRIB', 'PV114-R50-VT2']
PV_114_R50_SS = ['114-R50 SS', '114-R50 STX',
                 'PV 114 R50SS', 'PV 114-R50 -SSX']
PV_109_P29_VT2PRIB = ['109-P29 VT2', '109-P29VT2PRIB', 'PV 109-P29 - VT2', 'PV109-P29',
                      'PV109-P29 VT2', 'PV109-P29 VT2PRIB', 'PV109-P29 VT2PRIB 34k', 'PV109-P29 Vt2P', 'PV109-P29VT2', ]
PV_109_A90_VT2PRIB = ['109-A90', '109-A90 VT2', '109-A90 VT2P',
                      '109-A90 VT2PRIB', '109-A90VT2', '109-A90VT2PRIB', 'PV 109-A90 VT2', 'PV109-A90', 'PV109-A90VT2', 'PV 109-A90 VT2PRIB']
PV_3519X = ['3519X']
LG_5643_VT2P = ['LG5643 VT2', '5643 VT2',
                '5643VT2', '5643VT2RIB', 'LG5643 VT2', 'LG5643 VT2PRIB', 'LG5643 VT2PRIB']
LG_5643_STX = ['5643 STX', 'LG5643STX', '5643SS-RIB']
LG_5700_VT2P = ['LG5700VT2', '5700 VT2', '5700 VT2RIB',
                '5700VT2', 'LG5700 VT2', 'LG5700VT2', 'LG5700 VT2PRIB', 'LG5700 VT2PRIB']
LG_5700_STX = ['5700 STX', 'LG5700STX']
LG_5525 = ['LG5525']
LG_61C48_VT2P = ['LG61C48VT2P', 'LG61C48 VT2']
LG_62C35_VT2P = ['62C35 NT2', '62C35 VT2', '62C35VT2', '62C35VT2RIB',
                 'LG62C35', 'LG62C35 VT2', 'LG62C35VT2', 'LG62C35VT2P']
LG_64C30_TRC = ['LG64C30TRC', '64C30 VIP',
                '64C30 TRE', '64C30 VT2', '64C30TRC']
LG_59C66_VTP2 = ['59C66 VT2', 'LG59C66 VT2']
LG_59C41_STX = ['59C41 STX']
LG_66C32_STX = ['LG66C32STX', '66C32 STX',
                '66C32SSTX', 'LG66C32 STX', 'LG66C32STX']
LG_66C32_VT2P = ['66C32 VT2']
LG_66C28_3110 = ['66C28 - 3110', '66C28-3110', '66C28 3220', '66C28 VIP']
LG_67C45_STX = ['LG 67C45 STX RIB', '67C45 SS-RIB', '67C45STX', '67C45 STX']
LG_60C33_VT2P = ['LG60C33 VT2', '60C33 VT2']
LG_5650_VT2PRIB = ['LG5650 VT2PRIB']
LG_5606_STX = ['5606 STX']
LG_55C95_VT2P = ['55C95 VT2']
LG_58C77_VT2P = ['58C77 3220']
LG_59C72_VT2P = ['59C72 VT2']
LG_68C88_VT2P = ['68C88 VT2']
LG_C2888RX = ['2888RX']
LG_S2989RX = ['2989RX']
LG_S3060RX = ['3060RX']
LG_C3550RX = ['3550RX']
LG_S3600RX = ['3600RX']
LG_59C46 = ['59C46']
LG_64C30_TRC = ['64C30 TRE', '64C30TRC', 'LG64C30TRC',
                'LG6430 TRCIB', '64C30 VIP', '64C30 VT2']
MYCO_2410Q = ['2410 Q', '2410Q']
MYCO_2410AM = ['2410 AML', '2410 AM', '2410-AM',
               '2410AM', 'MY2410AM', 'MYCO 2410 AM']
MYCO_1610Q = ['1610', '1610 SS']
MYCO_1201Q = ['1201 Q', '1201 Qrome', '1201Q', 'MY1201Q']
MYCO_2470AM = ['MY2470AM']
MYCO_2470AML = ['2470 AML', '2470AML', 'MY2470AML', 'MYCO 2470 AML']
MYCO_2470Q = ['2470Q']
MYCO_MY1404AM = ['MY1404AM', 'MY1404 AM']
MYCO_1404AM = ['1404AM', '1404 AM']
MYCO_1830AML = []


def topPlantingInfo():
    GROWER_NAME = plantingSheet['C3'].value
    GROWER_CITY = plantingSheet['C4'].value
    COUNTY = plantingSheet['C5'].value
    ACE_LOCATION = plantingSheet['C6'].value
    if ACE_LOCATION != None:
        ACE_LOCATION = ACE_LOCATION
    else:
        ACE_LOCATION = "None"
    STARTED_PLOT_ON = plantingSheet['C7'].value
    FLAG_LOCATION = plantingSheet['C8'].value
    if FLAG_LOCATION != None:
        FLAG_LOCATION = FLAG_LOCATION
    else:
        FLAG_LOCATION = "None"
    GPS_LATITUDE = plantingSheet['C9'].value
    FUNGICIDE = plantingSheet['C10'].value
    if FUNGICIDE != None:
        FUNGICIDE = FUNGICIDE
    else:
        FUNGICIDE = "None"
    CROP = plantingSheet['H3'].value
    PLANTING_DATE = plantingSheet['H4'].value
    SEEDING_RATE = plantingSheet['H5'].value
    PLANTING_DEPTH_IN = plantingSheet['H6'].value
    if PLANTING_DEPTH_IN != None:
        PLANTING_DEPTH_IN = PLANTING_DEPTH_IN
    else:
        PLANTING_DEPTH_IN = "None"
    PLANTER_TYPE = plantingSheet['H7'].value
    if PLANTER_TYPE != None:
        PLANTER_TYPE = PLANTER_TYPE
    else:
        PLANTER_TYPE = "None"
    ROW_WIDTH = plantingSheet['H8'].value
    if ROW_WIDTH != None:
        ROW_WIDTH = ROW_WIDTH
    else:
        SOIL_TEXTURE = "None"
    GPS_LONGITUDE = plantingSheet['H9'].value
    HERBICIDE = plantingSheet['H10'].value
    if HERBICIDE != None:
        HERBICIDE = HERBICIDE
    else:
        HERBICIDE = "None"
    PLOT_TYPE = plantingSheet['L5'].value
    if PLOT_TYPE != None:
        PLOT_TYPE = PLOT_TYPE
    else:
        PLOT_TYPE = "None"
    IRRIGATION_TYPE = plantingSheet['L6'].value
    if IRRIGATION_TYPE != None:
        IRRIGATION_TYPE = IRRIGATION_TYPE
    else:
        IRRIGATION_TYPE = "None"
    PREVIOUS_CROP = plantingSheet['L7'].value
    if PREVIOUS_CROP != None:
        PREVIOUS_CROP = PREVIOUS_CROP
    else:
        PREVIOUS_CROP = "None"
    TILLAGE_SYSTEM = plantingSheet['L8'].value
    if TILLAGE_SYSTEM != None:
        TILLAGE_SYSTEM = TILLAGE_SYSTEM
    else:
        TILLAGE_SYSTEM = "None"
    SOIL_TEXTURE = plantingSheet['L9'].value
    if SOIL_TEXTURE != None:
        SOIL_TEXTURE = SOIL_TEXTURE
    else:
        SOIL_TEXTURE = "None"
    INSECTICIDE_RATE = plantingSheet['L10'].value
    if INSECTICIDE_RATE != None:
        INSECTICIDE_RATE = INSECTICIDE_RATE
    else:
        INSECTICIDE_RATE = "None"
    FORM_TYPE = "PLANTING FORM"

    f1.write(GROWER_NAME.title() + "\t" + GROWER_CITY.title() + "\t" + COUNTY.title() + "\t" + ACE_LOCATION.title() + "\t" + STARTED_PLOT_ON.title() + "\t" + FLAG_LOCATION.title() + "\t" + str(GPS_LATITUDE) + "\t" + FUNGICIDE.title() + "\t" + CROP.title() + "\t" + str(PLANTING_DATE) + "\t" + str(SEEDING_RATE) + "\t" +
             str(PLANTING_DEPTH_IN) + "\t" + PLANTER_TYPE.title() + "\t" + str(ROW_WIDTH) + "\t" + str(GPS_LONGITUDE) + "\t" + HERBICIDE.title() + "\t" + PLOT_TYPE.title() + "\t" + IRRIGATION_TYPE.title() + "\t" + PREVIOUS_CROP.title() + "\t" + TILLAGE_SYSTEM.title() + "\t" + SOIL_TEXTURE.title() + "\t" + INSECTICIDE_RATE + "\t" + FORM_TYPE.title() + "\t")


def bottomPlantingInfo():

    for row in range(17, plantingSheet.max_row + 1):
        ENTRY = plantingSheet['A' + str(row)].value
        COMPANY = plantingSheet['C' + str(row)].value
        BASEITEMGUID = ''
        HYBRID_VARIETY = plantingSheet['F' + str(row)].value
        for product in PV_113_V89:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'PV 113-V89 VT2PRIB'
                BASEITEMGUID = '8F552B16-DB63-48B0-8C99-DB6E968B1E22'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in NK_1082_5222A:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'NK 1082-5222A'
                BASEITEMGUID = '047B6471-E671-434C-AF2D-1DE5657F2AEA'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in PV_110_H20SS:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'PV 110-H20 SS'
                BASEITEMGUID = '22E1DB5A-F533-4CE2-BF66-153706205D5F'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in PV_110_H20VT2PRIB:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'PV 110-H20 VT2PRIB'
                BASEITEMGUID = '00B96AF0-0394-436A-A385-978A1C21BD89'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in PV_112_Q70BT2PRIB:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'PV 112-Q70 VT2PRIB'
                BASEITEMGUID = '608CB2B2-9219-414A-8486-EFC9C2542C86'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in PV_112_T69SS:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'PV 112-T69 SS RIB'
                BASEITEMGUID = '993A89E8-E3F7-47AE-9908-FE7CAAF9D6F7'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in PV_113_B40:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'PV 113-B40 DGVT2PRIB'
                BASEITEMGUID = 'B376985F-E07A-474F-8170-70B1A999C8B2'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in PV_115_D59:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'PV 115-D59 VT2PRIB'
                BASEITEMGUID = 'D27B9A20-13F0-4CFC-84EF-19ECE9B64924'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in PV_115_M60:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'PV 115-M60 TRECEPTA'
                BASEITEMGUID = 'F6B5F7F0-10CD-49D1-9192-57E84EE753EC'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in PV_114_R50_VT2PRIB:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'PV 114-R50 VT2PRIB'
                BASEITEMGUID = '9B16561A-D598-4F6F-BADA-AC91C027572B'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in PV_114_R50_SS:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'PV 114-R50 SS'
                BASEITEMGUID = '4B9042A1-9226-4BA1-A8E7-9AC8B16600C5'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in PV_109_P29_VT2PRIB:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'PV 109-P29 VT2PRIB'
                BASEITEMGUID = '94C559B2-C262-43B0-BDA3-E98F7EAC76CE'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in PV_109_A90_VT2PRIB:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'PV 109-A90 VT2PRIB'
                BASEITEMGUID = '3788CA80-FFEE-4898-9052-DFED4D0A1F3B'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in PV_3519X:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'PV 3519X'
                BASEITEMGUID = 'C88D2F86-62D5-476F-AE33-0D6603127C7B'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in LG_5643_VT2P:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'LG 5643 VT2P RIB'
                BASEITEMGUID = 'BD58062A-D5BC-4C71-9EB9-BABF52E8B384'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in LG_5643_STX:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'LG 5643 STX RIB'
                BASEITEMGUID = 'D524D3CE-82B8-4492-80AE-07663301546D'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in LG_5700_VT2P:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'LG 5700 VT2P RIB'
                BASEITEMGUID = '7B649F45-2512-474A-89FB-5D123DF4E1B6'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in LG_5700_STX:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'LG 5700 STX RIB'
                BASEITEMGUID = '4F6A0CAE-A6E7-4847-8361-D880BF29DE1C'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in LG_5525:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'LG 5525'
                BASEITEMGUID = '78AE89F4-83D0-49FD-B890-2A0AD3849B5E'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in LG_61C48_VT2P:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'LG 61C48 VT2P RIB'
                BASEITEMGUID = 'BE770144-C9C3-4EB7-85FD-3C3CC4A92532'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in LG_62C35_VT2P:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'LG 62C35 VT2P RIB'
                BASEITEMGUID = '597DE6C5-8C4A-4A3A-B89B-27BB5BAFAB96'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in LG_64C30_TRC:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'LG 64C30 TRC RIB'
                BASEITEMGUID = 'E4172F9A-F63B-448A-929D-5ADF65115C40'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in LG_59C66_VTP2:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'LG 59C66 VT2P RIB'
                BASEITEMGUID = '6531D4F6-9965-4A44-88E7-0E506A7DC5FC'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in LG_59C41_STX:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'LG 59C41 STX RIB'
                BASEITEMGUID = 'C76FE046-C980-48D3-BBD0-F8E84FE51D68'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in LG_66C32_STX:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'LG 66C32 STX RIB'
                BASEITEMGUID = 'F58A604D-0215-40B0-8ABD-6A2BA3D12F05'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in LG_66C32_VT2P:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'LG 66C32 VT2P RIB'
                BASEITEMGUID = 'D6CF601F-89C8-4663-A155-AA52903ED565'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in LG_66C28_3110:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'LG 66C28 3110'
                BASEITEMGUID = 'BBF34E36-114B-4C4C-8B5F-CB661D44A704'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in LG_67C45_STX:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'LG 67C45 STX RIB'
                BASEITEMGUID = 'DA404D4C-0325-46D9-8E6E-32262E50CB17'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in LG_60C33_VT2P:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'LG 60C33 VT2P RIB'
                BASEITEMGUID = '59767244-D7E8-4887-B141-B07F5A87269A'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in LG_5650_VT2PRIB:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'LG 5650 VT2P RIB'
                BASEITEMGUID = '5873E176-BC43-4B29-A144-7CC73F779996'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in LG_5606_STX:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'LG 5606 STX RIB'
                BASEITEMGUID = '9A73674F-E03D-4D07-970C-AC70FE0F4EB7'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in LG_55C95_VT2P:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'LG 55C95 VT2P'
                BASEITEMGUID = '89226E37-7A3B-46B1-B817-29D110347AA6'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in LG_58C77_VT2P:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'LG 58C77 VT2P RIB'
                BASEITEMGUID = 'A27CC89D-7177-41DC-AB41-FB77AE7E77F9'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in LG_59C72_VT2P:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'LG 59C72 VT2P RIB'
                BASEITEMGUID = '951D5426-254E-4582-B652-0D43F5E8D93D'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in LG_68C88_VT2P:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'LG 68C88 VT2P RIB'
                BASEITEMGUID = '29E7DE18-AC94-4ACF-89F5-A79F35974A77'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in LG_C2888RX:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'LG C2888RX'
                BASEITEMGUID = 'DBD6A924-13B6-4EEA-B30F-62D7CFD76924'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in LG_S2989RX:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'LG S2989RX'
                BASEITEMGUID = '65C6770F-EA57-403E-BDF6-8FDA37746E6D'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in LG_S3060RX:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'LG S3060RX'
                BASEITEMGUID = '5A2A26F5-0BD1-4778-91F9-4FF646B0AD6E'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in LG_C3550RX:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'LG C3550RX'
                BASEITEMGUID = '19A4C5B6-E434-4DD3-99BF-7DA444457779'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in LG_S3600RX:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'LG S3600RX'
                BASEITEMGUID = '0CC3D44A-E050-4336-B0D7-E83328067EA5'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in LG_59C46:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'LG 59C46 Conv'
                BASEITEMGUID = '4DE8DBBC-FD1D-46AE-A27F-97CDC7FA1084'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in LG_64C30_TRC:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'LG 64C30 TRC RIB'
                BASEITEMGUID = 'E4172F9A-F63B-448A-929D-5ADF65115C40'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in MYCO_2410Q:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'MYCO 2410Q'
                BASEITEMGUID = '15DCB22C-4426-4611-9118-270AD62B6799'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in MYCO_2410AM:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'MYCO 2410AM'
                BASEITEMGUID = '2E351241-17AC-4B98-BD1F-FA105E10B803'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in MYCO_1610Q:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'MYCO 1610Q'
                BASEITEMGUID = '878B3D71-D1D4-43C6-A1AA-F2CA6B22C7FF'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in MYCO_1201Q:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'MYCO 1201Q'
                BASEITEMGUID = '804D4E61-1A96-47CB-9559-F298FC4D7E2C'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in MYCO_2470AM:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'MYCO 2470AM'
                BASEITEMGUID = '4EB2A823-A071-4A7E-97AF-76F269730B9B'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in MYCO_2470AML:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'MYCO 2470AML'
                BASEITEMGUID = 'ADDFC5AB-B968-4B57-8240-D8A4AFFEA2EA'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in MYCO_2470Q:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'MYCO 2470Q'
                BASEITEMGUID = '39671E4E-80D4-4CB6-B9B9-D16A0DBE542F'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in MYCO_MY1404AM:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'MYCO MY1404AM'
                BASEITEMGUID = 'B5B2A4AA-1B53-447F-A93F-5E79B1275DA6'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in MYCO_1404AM:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'MYCO 1404AM'
                BASEITEMGUID = '378F6F09-0882-41C6-9AC7-78B6D662CB89'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in MYCO_1830AML:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'MYCO 1830AML'
                BASEITEMGUID = '5EB16F04-6313-4EBC-8DD0-1E715E48D7CB'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY

        SEED_TREATMENTS = plantingSheet['J' + str(row)].value
        if SEED_TREATMENTS != None:
            SEED_TREATMENTS = SEED_TREATMENTS
        else:
            SEED_TREATMENTS = "None"
        NUM_OF_ROWS = plantingSheet['M' + str(row)].value

        if COMPANY != None and HYBRID_VARIETY != None:
            topPlantingInfo()
            f1.write(str(ENTRY) + "\t" + COMPANY.title() + "\t" + str(HYBRID_VARIETY) +
                     "\t" + str(BASEITEMGUID) +
                     "\t" + str(SEED_TREATMENTS) + "\t" + str(NUM_OF_ROWS) + "\n")
        else:
            continue


for folders, sub_folders, file in os.walk(docDir):
    for name in file:
        if name.endswith(".xlsx"):
            filename = os.path.join(folders, name)
            print(filename)
            wb = load_workbook(filename)
            plantingSheet = wb['PLANTING FORM']
            with open(r'C:\Users\rkeenan\OneDrive - Aurora Cooperative\Documents\2020 Tableau Updates\Al Perry\Test Plot(PLANTING).txt', 'a') as f1:
                bottomPlantingInfo()
                wb.close()
        else:
            continue

print("Your Planting Form data plot file is done!")

# with open(r'C:\Users\rkeenan\OneDrive - Aurora Cooperative\Documents\2020 Tableau Updates\Al Perry\Test Plot(PLANTING).txt', 'a+') as f1:
#    for line in f1:
#        if '113-V89' in line:
#            line.replace('113-V89', 'PV 113-V89 VT2PRIB')
#        else:
#            continue

# print("Product names are now updated.")
# wb.close()


# bottomPlantingInfo()
# plantingWriteFile.close()
# wb.close()

# This is for the NOTES FORM

# wb = openpyxl.load_workbook(
#    r'C:\Users\rkeenan\OneDrive - Aurora Cooperative\Documents\2020 Tableau Updates\Al Perry\2020 Aaron Franson Test Plot.xlsx')
# notesSheet = wb['NOTES FORM']
# notesWriteFile = open(
#    r'C:\Users\rkeenan\OneDrive - Aurora Cooperative\Documents\2020 Tableau Updates\Al Perry\Test Plot(NOTES).txt', 'w')


def topNotesInfo():
    GROWER_NAME = notesSheet['C3'].value
    GROWER_CITY = notesSheet['C4'].value
    COUNTY = notesSheet['C5'].value
    ACE_LOCATION = notesSheet['C6'].value
    STATED_PLOT_ON = notesSheet['C7'].value
    FLAT_LOCATION = notesSheet['C8'].value
    GPS_LATITUDE = notesSheet['C9'].value
    FUNGICIDE = notesSheet['C10'].value
    if FUNGICIDE != None:
        FUNGICIDE = FUNGICIDE
    else:
        FUNGICIDE = "None"
    CROP = notesSheet['H3'].value
    PLANTING_DATE = notesSheet['H4'].value
    SEEDING_RATE = notesSheet['H5'].value
    PLANTING_DEPTH_IN = notesSheet['H6'].value
    PLANTER_TYPE = notesSheet['H7'].value
    ROW_WIDTH = notesSheet['H8'].value
    GPS_LONGITUDE = notesSheet['H9'].value
    HERBICIDE = notesSheet['H10'].value
    if HERBICIDE != None:
        HERBICIDE = HERBICIDE
    else:
        HERBICIDE = "None"
    PLOT_TYPE = notesSheet['L5'].value
    IRRIGATION_TYPE = notesSheet['L6'].value
    PREVIOUS_CROP = notesSheet['L7'].value
    TILLAGE_SYSTEM = notesSheet['L8'].value
    SOIL_TEXTURE = notesSheet['L9'].value
    INSECTICIDE_RATE = notesSheet['L10'].value
    if INSECTICIDE_RATE != None:
        INSECTICIDE_RATE = INSECTICIDE_RATE
    else:
        INSECTICIDE_RATE = "None"
    FORM_TYPE = "NOTES FORM"

    notesWriteFile.write(GROWER_NAME + "\t" + GROWER_CITY + "\t" + COUNTY + "\t" + ACE_LOCATION + "\t" + STATED_PLOT_ON + "\t" + FLAT_LOCATION + "\t" + str(GPS_LATITUDE) + "\t" + FUNGICIDE + "\t" + CROP + "\t" + str(PLANTING_DATE) + "\t" + str(SEEDING_RATE) + "\t" +
                         str(PLANTING_DEPTH_IN) + "\t" + PLANTER_TYPE + "\t" + str(ROW_WIDTH) + "\t" + str(GPS_LONGITUDE) + "\t" + HERBICIDE + "\t" + PLOT_TYPE + "\t" + IRRIGATION_TYPE + "\t" + PREVIOUS_CROP + "\t" + TILLAGE_SYSTEM + "\t" + SOIL_TEXTURE + "\t" + INSECTICIDE_RATE + "\t" + FORM_TYPE + "\t")


def bottomNotesInfo():
    for row in range(15, notesSheet.max_row + 1):
        HYBRID_VARIETY = notesSheet['A' + str(row)].value
        EMERGENCE = notesSheet['C' + str(row)].value
        if EMERGENCE != None:
            EMERGENCE = EMERGENCE
        else:
            EMERGENCE = "None"
        EARLY_VIGOR = notesSheet['D' + str(row)].value
        if EARLY_VIGOR != None:
            EARLY_VIGOR = EARLY_VIGOR
        else:
            EARLY_VIGOR = "None"
        LATE_PLANT_HEALTH = notesSheet['E' + str(row)].value
        if LATE_PLANT_HEALTH != None:
            LATE_PLANT_HEALTH = LATE_PLANT_HEALTH
        else:
            LATE_PLANT_HEALTH = "None"
        GENERAL_COMMENTS = notesSheet['F' + str(row)].value
        if GENERAL_COMMENTS != None:
            GENERAL_COMMENTS = GENERAL_COMMENTS
        else:
            GENERAL_COMMENTS = "None"
        ENTRY = notesSheet['M' + str(row)].value

        if HYBRID_VARIETY != None and EMERGENCE != None:
            topNotesInfo()
            notesWriteFile.write(str(HYBRID_VARIETY) + "\t" + str(EMERGENCE) + "\t" + str(EARLY_VIGOR) +
                                 "\t" + str(LATE_PLANT_HEALTH) + "\t" + str(GENERAL_COMMENTS) + "\t" + str(GENERAL_COMMENTS) + "\n")
        else:
            continue

    print("Your Notes Form data plot file is done!")


# bottomNotesInfo()
# notesWriteFile.close()


# This is for the HARVEST FORM

# harvestSheet = wb['HARVEST FORM']
# harvestWriteFile = open(
#    r'C:\Users\rkeenan\OneDrive - Aurora Cooperative\Documents\2020 Tableau Updates\Al Perry\Test Plot(HARVEST).txt', 'w')

def topHarvestInfo():
    GROWER_NAME = harvestSheet['C3'].value
    GROWER_CITY = harvestSheet['C4'].value
    COUNTY = harvestSheet['C5'].value
    ACE_LOCATION = harvestSheet['C6'].value
    STATED_PLOT_ON = harvestSheet['C7'].value
    FLAT_LOCATION = harvestSheet['C8'].value
    GPS_LATITUDE = harvestSheet['C9'].value
    FUNGICIDE = harvestSheet['C10'].value
    if FUNGICIDE != None:
        FUNGICIDE = FUNGICIDE
    else:
        FUNGICIDE = "None"
    HARVEST_DATE = harvestSheet['C11'].value
    if HARVEST_DATE != None:
        HARVEST_DATE = HARVEST_DATE
    else:
        HARVEST_DATE = "None"
    CROP = harvestSheet['H3'].value
    PLANTING_DATE = harvestSheet['H4'].value
    SEEDING_RATE = harvestSheet['H5'].value
    PLANTING_DEPTH_IN = harvestSheet['H6'].value
    PLANTER_TYPE = harvestSheet['H7'].value
    ROW_WIDTH = harvestSheet['H8'].value
    GPS_LONGITUDE = harvestSheet['H9'].value
    HERBICIDE = harvestSheet['H10'].value
    if HERBICIDE != None:
        HERBICIDE = HERBICIDE
    else:
        HERBICIDE = "None"
    COMMODITY_PRICE = harvestSheet['H11'].value
    PLOT_TYPE = harvestSheet['L5'].value
    IRRIGATION_TYPE = harvestSheet['L6'].value
    PREVIOUS_CROP = harvestSheet['L7'].value
    TILLAGE_SYSTEM = harvestSheet['L8'].value
    SOIL_TEXTURE = harvestSheet['L9'].value
    INSECTICIDE_RATE = harvestSheet['L10'].value
    if INSECTICIDE_RATE != None:
        INSECTICIDE_RATE = INSECTICIDE_RATE
    else:
        INSECTICIDE_RATE = "None"
    DRYING_COST = harvestSheet['L11'].value
    FORM_TYPE = "HARVEST FORM"

    harvestWriteFile.write(GROWER_NAME + "\t" + GROWER_CITY + "\t" + COUNTY + "\t" + ACE_LOCATION + "\t" + STATED_PLOT_ON + "\t" + FLAT_LOCATION + "\t" + str(GPS_LATITUDE) + "\t" + FUNGICIDE + "\t" + HARVEST_DATE + "\t" + CROP + "\t" + str(PLANTING_DATE) + "\t" + str(SEEDING_RATE) + "\t" +
                           str(PLANTING_DEPTH_IN) + "\t" + PLANTER_TYPE + "\t" + str(ROW_WIDTH) + "\t" + str(GPS_LONGITUDE) + "\t" + HERBICIDE + "\t" + str(COMMODITY_PRICE) + "\t" + PLOT_TYPE + "\t" + IRRIGATION_TYPE + "\t" + PREVIOUS_CROP + "\t" + TILLAGE_SYSTEM + "\t" + SOIL_TEXTURE + "\t" + INSECTICIDE_RATE + "\t" + str(DRYING_COST) + "\t" + FORM_TYPE + "\t")


def bottomHarvestInfo():
    for row in range(17, harvestSheet.max_row + 1):
        BRAND = harvestSheet['A' + str(row)].value
        PRODUCT = harvestSheet['C' + str(row)].value
        ROW_LENGTH = harvestSheet['F' + str(row)].value
        if ROW_LENGTH != None:
            ROW_LENGTH = ROW_LENGTH
        else:
            ROW_LENGTH = "None"
        WET_WEIGHT = harvestSheet['G' + str(row)].value
        if WET_WEIGHT != None:
            WET_WEIGHT = WET_WEIGHT
        else:
            WET_WEIGHT = "None"
        HARVEST_MOISTURE_PCT = harvestSheet['H' + str(row)].value
        if HARVEST_MOISTURE_PCT != None:
            HARVEST_MOISTURE_PCT = HARVEST_MOISTURE_PCT
        else:
            HARVEST_MOISTURE_PCT = "None"
        NUM_OF_ROWS = harvestSheet['I' + str(row)].value
        TEST_WEIGHT = harvestSheet['J' + str(row)].value
        if TEST_WEIGHT != None:
            TEST_WEIGHT = TEST_WEIGHT
        else:
            TEST_WEIGHT = "None"
        YIELD_PER_ACRE = harvestSheet['K' + str(row)].value
        if YIELD_PER_ACRE != None:
            YIELD_PER_ACRE = YIELD_PER_ACRE
        else:
            YIELD_PER_ACRE = "None"
        PCT_OF_PLOT_ADVANTAGE = harvestSheet['L' + str(row)].value
        if PCT_OF_PLOT_ADVANTAGE != None:
            PCT_OF_PLOT_ADVANTAGE = PCT_OF_PLOT_ADVANTAGE
        else:
            PCT_OF_PLOT_ADVANTAGE = "None"
        YIELD_PER_ACRE_RANK = harvestSheet['M' + str(row)].value
        if YIELD_PER_ACRE_RANK != None:
            YIELD_PER_ACRE_RANK = YIELD_PER_ACRE_RANK
        else:
            YIELD_PER_ACRE_RANK = "None"
        DOLLARS_PER_ACRE_RANK = harvestSheet['N' + str(row)].value
        if DOLLARS_PER_ACRE_RANK != None:
            DOLLARS_PER_ACRE_RANK = DOLLARS_PER_ACRE_RANK
        else:
            DOLLARS_PER_ACRE_RANK = "None"
        if BRAND != None and PRODUCT != None:
            topHarvestInfo()
            harvestWriteFile.write(BRAND + "\t" + str(PRODUCT) + "\t" + str(ROW_LENGTH) +
                                   "\t" + str(WET_WEIGHT) + "\t" + str(HARVEST_MOISTURE_PCT) + "\t" + str(NUM_OF_ROWS) + "\t" + str(TEST_WEIGHT) + "\t" + str(YIELD_PER_ACRE) + "\t" + str(PCT_OF_PLOT_ADVANTAGE) + "\t" + str(YIELD_PER_ACRE_RANK) + "\t" + str(DOLLARS_PER_ACRE_RANK) + "\n")
        else:
            continue

    print("Your Harvest Form data plot file is done!")

# bottomHarvestInfo()
# harvestWriteFile.close()
