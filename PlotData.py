import openpyxl
import pprint

wb = openpyxl.load_workbook(
    r'C:\Users\rkeenan\OneDrive - Aurora Cooperative\Documents\2020 Tableau Updates\Al Perry\2020 Aaron Franson Test Plot.xlsx')
sheet = wb['PLANTING FORM']
writeFile = open(
    r'C:\Users\rkeenan\OneDrive - Aurora Cooperative\Documents\2020 Tableau Updates\Al Perry\2020 Aaron Franson Test Plot.txt', 'w')


def topPlantingInfo():
    GROWER_NAME = sheet['C3'].value
    GROWER_CITY = sheet['C4'].value
    COUNTY = sheet['C5'].value
    ACE_LOCATION = sheet['C6'].value
    STATED_PLOT_ON = sheet['C7'].value
    FLAT_LOCATION = sheet['C8'].value
    GPS_LATITUDE = sheet['C9'].value
    FUNGICIDE = sheet['C10'].value
    if FUNGICIDE != None:
        FUNGICIDE = FUNGICIDE
    else:
        FUNGICIDE = "None"
    CROP = sheet['H3'].value
    PLANTING_DATE = sheet['H4'].value
    SEEDING_RATE = sheet['H5'].value
    PLANTING_DEPTH_IN = sheet['H6'].value
    PLANTER_TYPE = sheet['H7'].value
    ROW_WIDTH = sheet['H8'].value
    GPS_LONGITUDE = sheet['H9'].value
    HERBICIDE = sheet['H10'].value
    if HERBICIDE != None:
        HERBICIDE = HERBICIDE
    else:
        HERBICIDE = "None"
    PLOT_TYPE = sheet['L5'].value
    IRRIGATION_TYPE = sheet['L6'].value
    PREVIOUS_CROP = sheet['L7'].value
    TILLAGE_SYSTEM = sheet['L8'].value
    SOIL_TEXTURE = sheet['L9'].value
    INSECTICIDE_RATE = sheet['L10'].value
    if INSECTICIDE_RATE != None:
        INSECTICIDE_RATE = INSECTICIDE_RATE
    else:
        INSECTICIDE_RATE = "None"
    FORM_TYPE = "PLANTING FORM"

    writeFile.write(GROWER_NAME + "\t" + GROWER_CITY + "\t" + COUNTY + "\t" + ACE_LOCATION + "\t" + STATED_PLOT_ON + "\t" + FLAT_LOCATION + "\t" + str(GPS_LATITUDE) + "\t" + FUNGICIDE + "\t" + CROP + "\t" + str(PLANTING_DATE) + "\t" + str(SEEDING_RATE) + "\t" +
                    str(PLANTING_DEPTH_IN) + "\t" + PLANTER_TYPE + "\t" + str(ROW_WIDTH) + "\t" + str(GPS_LONGITUDE) + "\t" + HERBICIDE + "\t" + PLOT_TYPE + "\t" + IRRIGATION_TYPE + "\t" + PREVIOUS_CROP + "\t" + TILLAGE_SYSTEM + "\t" + SOIL_TEXTURE + "\t" + INSECTICIDE_RATE + "\t" + FORM_TYPE + "\t")


def bottomPlantingInfo():
    for row in range(17, sheet.max_row + 1):
        ENTRY = sheet['A' + str(row)].value
        COMPANY = sheet['C' + str(row)].value
        HYBRID_VARIETY = sheet['F' + str(row)].value
        SEED_TREATMENTS = sheet['J' + str(row)].value
        if SEED_TREATMENTS != None:
            SEED_TREATMENTS = SEED_TREATMENTS
        else:
            SEED_TREATMENTS = "None"
        NUM_OF_ROWS = sheet['M' + str(row)].value

        if COMPANY != None and HYBRID_VARIETY != None:
            topPlantingInfo()
            writeFile.write(str(ENTRY) + "\t" + COMPANY + "\t" + str(HYBRID_VARIETY) +
                            "\t" + SEED_TREATMENTS + "\t" + str(NUM_OF_ROWS) + "\n")
        else:
            continue

# def topNotesInfo():

# def bottomNotesInfo():

# def topHarvestInfo():

# def bottomHarvestInfo():


bottomPlantingInfo()
writeFile.close()
