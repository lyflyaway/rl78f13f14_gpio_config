import xlrd
import time

COLUMN_NUM = 113
PORT_NUM = 14

x1 = xlrd.open_workbook("rl78f13f14_gpio.xlsx")
sheet1 = x1.sheet_by_name("sleep config")

column0 = sheet1.col_values(1)
column1 = sheet1.col_values(2)

if (len(column0) == COLUMN_NUM) and (len(column1) == COLUMN_NUM):
    dir = [0x00] * PORT_NUM
    lev = [0x00] * PORT_NUM
    pup = [0x00] * PORT_NUM
    error = 0
    j = 0
    for i in range(1, COLUMN_NUM):
        if (i != 1) and ((i-1) % 8 == 0):
            j = j + 1
        
        dir[j] >>= 1
        lev[j] >>= 1
        pup[j] >>= 1
            
        if (column0[i] == "input") or (column0[i] == "-"):
            dir[j] |= 0x80

        elif (column0[i] == "output low"):
            pass
            
        elif (column0[i] == "output high"):
            lev[j] |= 0x80
        
        else:
            print('column0[%d] error!' % i)
            error = 1
            break
            
            
        if (column1[i] == "no") or (column1[i] == "-"):
            pass
            
        elif column1[i] == "yes":
            pup[j] |= 0x80
            
        else:
            print('column1[%d] error!' % i)
            error = 1
            break
    
    if error == 0:
        #for j in range(0, PORT_NUM):
        #    print('0x%02X, 0x%02X, 0x%02X' % (dir[j], lev[j], pup[j]))
        config = "#ifndef GPIO_CONFIG_H__\n\
#define GPIO_CONFIG_H__\n\
\n\
typedef struct\n\
{\n\
    unsigned char pm;   //port mode\n\
    unsigned char p;    //port level\n\
    unsigned char pu;   //pull up\n\
}PortRegTy;\n\
\n\
const PortRegTy P0SleepConfig  = { 0x%02X, 0x%02X, 0x%02X };\n\
const PortRegTy P1SleepConfig  = { 0x%02X, 0x%02X, 0x%02X };\n\
//has no Port2\n\
const PortRegTy P3SleepConfig  = { 0x%02X, 0x%02X, 0x%02X };\n\
const PortRegTy P4SleepConfig  = { 0x%02X, 0x%02X, 0x%02X };\n\
const PortRegTy P5SleepConfig  = { 0x%02X, 0x%02X, 0x%02X };\n\
const PortRegTy P6SleepConfig  = { 0x%02X, 0x%02X, 0x%02X };\n\
const PortRegTy P7SleepConfig  = { 0x%02X, 0x%02X, 0x%02X };\n\
const PortRegTy P8SleepConfig  = { 0x%02X, 0x%02X, 0x%02X };\n\
const PortRegTy P9SleepConfig  = { 0x%02X, 0x%02X, 0x%02X };\n\
const PortRegTy P10SleepConfig = { 0x%02X, 0x%02X, 0x%02X };\n\
//has no Port11\n\
const PortRegTy P12SleepConfig = { 0x%02X, 0x%02X, 0x%02X };\n\
const PortRegTy P13SleepConfig = { 0x%02X, 0x%02X, 0x%02X };\n\
const PortRegTy P14SleepConfig = { 0x%02X, 0x%02X, 0x%02X };\n\
const PortRegTy P15SleepConfig = { 0x%02X, 0x%02X, 0x%02X };\n\
\n\
#endif/* GPIO_CONFIG_H__ */" % (dir[0], lev[0], pup[0],\
        dir[1], lev[1], pup[1],\
        dir[2], lev[2], pup[2],\
        dir[3], lev[3], pup[3],\
        dir[4], lev[4], pup[4],\
        dir[5], lev[5], pup[5],\
        dir[6], lev[6], pup[6],\
        dir[7], lev[7], pup[7],\
        dir[8], lev[8], pup[8],\
        dir[9], lev[9], pup[9],\
        dir[10], lev[10], pup[10],\
        dir[11], lev[11], pup[11],\
        dir[12], lev[12], pup[12],\
        dir[13], lev[13], pup[13])

        fp = open("gpio_config.h", "w+")
        fp.write(config)
        fp.close()
        print('Configure successfully.')
    
else:
    print('column = %d, %d error!' % (len(column0), len(column1)))
    
time.sleep(5)
    