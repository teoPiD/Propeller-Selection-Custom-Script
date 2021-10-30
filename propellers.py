import xlwt

LBF_TO_KG = 0.45359237
HP_TO_W = 745.699872 #assumed mechanical horsepower

APCData = "APC data/STATIC-2.dat"
outputFileName = "propellers"

counter = 0

countPropellers = 0
countTotalPropellers = 0

hoveringThrust = 1.0
maximumThrust = 2.0

printData = False
printNext = False

#In mode 0 we are looking for hoveringThrust
#and in mode 1 we are looking for maximumThrust
#in mode 2 we don't need to look any further
thrustLooker = 0

if __name__ == '__main__':
    
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet("Propeller Selection")
    greenBackground = xlwt.easyxf('pattern: pattern solid, fore_colour green;') 
  
    with open(APCData) as file:
        #skip first two lines
        next(file)
        next(file)
        for line in file:
            counter += 1
            if (line == '\n'):
                if(counter > 5):
                    counter = 0
                else:
                    counter -= 1
                continue
            elif (counter == 1):
                name = line.strip().split()[0]
                diameter = int(name[:name.index('x')])
                printData = False
                printNext = False
                dataHover = []
                if diameter >= 8000:
                    diameter /= 1000
                if diameter >= 800:
                    diameter /= 100
                elif diameter >= 80:
                    diameter /= 10
                if(diameter >= 8 and diameter < 11):
                    printData = True
                thrustLooker = 0
                countTotalPropellers += 1
            elif (counter > 3 and printData):
                line = line.strip().split()
                thrust = float(line[1])*LBF_TO_KG
                if(thrustLooker == 0):
                    if(thrust >= hoveringThrust):
                        power = float(line[2]) * HP_TO_W
                        gByW = thrust * 1000 / power
                        if gByW > 7.0:
                            dataHover = [line[0], thrust, power, gByW]
                            printNext = True
                        thrustLooker = 1
                elif(thrustLooker == 1 and printNext):
                    if(thrust >= maximumThrust):
                        power = float(line[2]) * HP_TO_W
                        gByW = thrust * 1000 / power
                        if(gByW > 4):
                            sheet.write(5*(countPropellers//2), 2 + ((countPropellers%2)*6), name) 
                            sheet.write(5*(countPropellers//2)+1, 1 + ((countPropellers%2)*6), 'RPM')
                            sheet.write(5*(countPropellers//2)+1, 2 + ((countPropellers%2)*6), 'Thrust')
                            sheet.write(5*(countPropellers//2)+1, 3 + ((countPropellers%2)*6), 'Power')
                            sheet.write(5*(countPropellers//2)+1, 4 + ((countPropellers%2)*6), 'gByW')
                            sheet.write(5*(countPropellers//2)+2, 0 + ((countPropellers%2)*6), 'Hover')
                            sheet.write(5*(countPropellers//2)+2, 1 + ((countPropellers%2)*6), dataHover[0])
                            sheet.write(5*(countPropellers//2)+2, 2 + ((countPropellers%2)*6), dataHover[1])
                            sheet.write(5*(countPropellers//2)+2, 3 + ((countPropellers%2)*6), dataHover[2])
                            if dataHover[3] >= 7:
                                sheet.write(5*(countPropellers//2)+2, 4 + ((countPropellers%2)*6), dataHover[3], greenBackground)
                            else:
                                sheet.write(5*(countPropellers//2)+2, 4 + ((countPropellers%2)*6), dataHover[3])
                            sheet.write(5*(countPropellers//2)+3, 0 + ((countPropellers%2)*6), 'Maximum')
                            sheet.write(5*(countPropellers//2)+3, 1 + ((countPropellers%2)*6), line[0])
                            sheet.write(5*(countPropellers//2)+3, 2 + ((countPropellers%2)*6), thrust)
                            sheet.write(5*(countPropellers//2)+3, 3 + ((countPropellers%2)*6), power)
                        
                            if gByW >= 5:
                                sheet.write(5*(countPropellers//2)+3, 4 + ((countPropellers%2)*6), gByW, greenBackground)
                            else:
                                sheet.write(5*(countPropellers//2)+3, 4 + ((countPropellers%2)*6), gByW)
                            #print("Maximum", line[0], thrust, power, gByW)

                            countPropellers += 1
                        thrustLooker = 2
                        
    print("Total propellers found that correspond to the requirements is %d in a total of %d propellers" %(countPropellers, countTotalPropellers))
    workbook.save(outputFileName + ".xls") 
