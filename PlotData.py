import os
import pprint
#import openpyxl
from openpyxl import load_workbook

# This is for the PLANTING FORM

docDir = r'C:\Users\rkeenan\OneDrive - Aurora Cooperative\Documents\2020 Tableau Updates\Al Perry'

plantingWriteFile = (
    r'C:\Users\rkeenan\OneDrive - Aurora Cooperative\Documents\2020 Tableau Updates\Al Perry\Test Plot(PLANTING).txt', 'w')

PV113_V89 = ['113-V89', '113-V89 VT2', '113-V89 VT2P', '113-V89 VT2PRIB', '113-V89VT2', '113-V89VT2PRIB', '113V89 VT2', 'PV 113 V89 VT2', 'PV 113-V89',
             'PV 113-V89 - VT2', 'PV 113-V89 VT2PRIB 32k', 'PV113 V89 VT2PRIB', 'PV113-V89', 'PV113-V89 VT2', 'PV113-V89 VT2RIB- Check', 'PV113-V89-VT2PRIB', 'PV113-V89VT2']

NK1082_5222A = ['1082 5222', '1082-5222', '1082-5222A', '1082A 5222',
                'NK1082-5222A', 'NK1082-5222A Brand', 'NK1082-5222AEZ1']

PV110_H20SS = ['110-H20 SS']

PV110_H20VT2PRIB = ['110-H20 VT2', '110-H20 VT2P', '110-H20 VT2PRIB', '110-H20VT2', '110-H20VT2PRIB', 'PV110 H20 VT2PRIB',
                    'PV110-H20 VT2', 'PV110-H20 VT2PRIB', 'PV110-H20 VT2PRIB 34k', 'PV110-H20 VT2RIB', 'PV110-H20VT2', ]
PV112_Q70BT2PRIB = ['112-Q70', '112-Q70 VT2',
                    '112-Q70 VT2P', '112-Q70 VT2PRIB', '112-Q70VT2', 'PV 112- Q70', 'PV 112-Q70 - VT2', 'PV 112-Q70 VT2', 'PV112 Q70 VT2PRIB', 'PV112-Q70', 'PV112-Q70 VT2', 'PV112-Q70 VT2PRIB', 'PV112-Q70 VT2PRIB 36k', 'PV112-Q70TT2', 'PV112Q-70']

PV112_T69SS = ['112-T69 SS', '112-T69 SSTX',
               '112-T69 SSTX RIB', '112-T69 STX', '112-T69SSRIB', 'PV 112-T69 - VT2', 'PV112 T69 SSRIB', 'PV112-T69', 'PV112-T69 SS', 'PV112-T69 SSRIB', 'PV112-T69 SSRIB 36k']
PV113_B40 = ['113-B40 DGVT2', '113-B40 DGVT2PRIB',
             '113-B40 VT2', '113-B40 VT2P', '113-B40DGVT2', 'PV113 B40 DGVT2PRIB', 'PV113-B40', 'PV113-B40 DGVT2PRIB', 'PV113-B40 VT2', 'PV113-B40DGVT2']


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
        for product in PV113_V89:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'PV 113-V89 VT2PRIB'
                BASEITEMGUID = '8F552B16-DB63-48B0-8C99-DB6E968B1E22'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in NK1082_5222A:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'NK 1082-5222A'
                BASEITEMGUID = '047B6471-E671-434C-AF2D-1DE5657F2AEA'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in PV110_H20SS:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'PV 110-H20 SS'
                BASEITEMGUID = '22E1DB5A-F533-4CE2-BF66-153706205D5F'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in PV110_H20VT2PRIB:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'PV 110-H20 VT2PRIB'
                BASEITEMGUID = '00B96AF0-0394-436A-A385-978A1C21BD89'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in PV112_Q70BT2PRIB:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'PV 112-Q70 VT2PRIB'
                BASEITEMGUID = '608CB2B2-9219-414A-8486-EFC9C2542C86'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in PV112_T69SS:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'PV 112-T69 SS RIB'
                BASEITEMGUID = '993A89E8-E3F7-47AE-9908-FE7CAAF9D6F7'
            else:
                HYBRID_VARIETY = HYBRID_VARIETY
        for product in PV113_B40:
            if HYBRID_VARIETY == product:
                HYBRID_VARIETY = 'PV 113-B40 DGVT2PRIB'
                BASEITEMGUID = 'B376985F-E07A-474F-8170-70B1A999C8B2'
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

#print("Product names are now updated.")
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
