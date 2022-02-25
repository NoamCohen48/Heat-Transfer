import xlsxwriter

# - - - - - - - משתנים - - - - - - -
Dt = 0.01      # דיוק אוילר
SectionAmount = 100     # כמות החתיכות

InOven = 0.08       # כמה מטרים בתוך התנור
Length = 0.65        # האורך הכולל של המוט
Diameter = 0.0118   # קוטר המוט
Mass = 576     # מסה בגרמים

K = 100
SpecificHeatCapacity = 0.9     # קיבול החום ביחידות J / (g * k)

M = 10.15     # מקדם M שמחושב בקובץ אקסל

Toven = 65.4
Tair = 23.7

Untill = 60     # time in minutes
NameExcel = 'Iron'
ExportTimeStep = 60     # שניות

# - - - - - - - פונקציות - - - - - - -

def PI():
    return 3.14159265359

def ToCelsius(Kelvin):
    return Kelvin - 273.15


def ToKelvin(Celsius):
    return Celsius + 273.15


# - - - - - - - פרמטרים - - - - - - -

SectionLength = (Length - InOven) / SectionAmount
SectionMass = Mass / SectionAmount
Radius = Diameter / 2
Area = PI() * Radius ** 2
Perimeter = Diameter * PI()
Untill = Untill * 60
H = M**2 * K * Area / Perimeter

time = 0
ExportCount = 0
# - - - - - - - יצירה - - - - - - -

Temp = []
Q = []

for i in range(SectionAmount):
    Temp.append(0)
    Q.append(0)

# - - - - - - - יצוא לאקסל - - - - - - -
workbook = xlsxwriter.Workbook(NameExcel + '.xlsx')
worksheet = workbook.add_worksheet()

# - - - - - - - לולאה - - - - - - -

worksheet.write(0, 0, 'Length = {}'.format(Length))
worksheet.write(0, 1, 'InOven = {}'.format(InOven))
worksheet.write(0, 2, 'Diameter = {}'.format(Diameter))
worksheet.write(0, 3, 'Mass = {}'.format(Mass))
worksheet.write(0, 4, 'C = {}'.format(SpecificHeatCapacity))
worksheet.write(0, 5, 'K = {}'.format(K))
worksheet.write(0, 6, 'H = {}'.format(H))
worksheet.write(0, 7, 'Msq = {}'.format((H * Perimeter)/(K * Area)))
worksheet.write(0, 8, 'Toven = {}'.format(Toven))
worksheet.write(0, 9, 'Tair = {}'.format(Tair))

worksheet.write(2 + ExportCount, 1, ExportCount)
for SectionNum in range(1, SectionAmount):
    worksheet.write(1, 1 + SectionNum, (SectionNum - 0.5) * SectionLength)
    worksheet.write(2 + ExportCount, 1 + SectionNum, 0)
ExportCount += 1

while time < Untill:

    q = K * (Temp[0] - (Toven - Tair)) * InOven * Perimeter
    Q[0] -= q * Dt
    for SectionNum in range(1, SectionAmount):
        # בתוך המוט
        q = K * (Temp[SectionNum - 1] - Temp[SectionNum]) * Area / SectionLength
        Q[SectionNum - 1] -= q * Dt
        Q[SectionNum] += q * Dt

        # עם הסביבה
        q = H * Temp[SectionNum] * SectionLength * Perimeter
        Q[SectionNum] -= q * Dt

    Temp[0] += Q[0] / (SpecificHeatCapacity * InOven / SectionLength * SectionMass)
    Q[0] = 0
    for SectionNum in range(1, SectionAmount):
        Temp[SectionNum] += Q[SectionNum] / (SpecificHeatCapacity * SectionMass)
        Q[SectionNum] = 0

    if ExportCount * ExportTimeStep <= time:
        worksheet.write(2 + ExportCount, 1, ExportCount)
        for SectionNum in range(1, SectionAmount):
            worksheet.write(2 + ExportCount, 1 + SectionNum, Temp[SectionNum])
        ExportCount += 1

    time += Dt

workbook.close()

print('finished')