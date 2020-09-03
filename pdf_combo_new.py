import enum
import glob
import openpyxl
import os
import pytesseract
import re
import time
import shutil
import string
import sys
import tabula
import warnings
import numpy  as np
import pandas as pd
import cv2
from os        import system, name
from PIL       import Image, ImageEnhance, ImageFilter
from PyPDF2    import PdfFileReader, PdfFileWriter, PdfFileMerger
from pdf2image import convert_from_path

warnings.filterwarnings("ignore")

# update this path - copy / paste file location after the r' below:
#pytesseract.pytesseract.tesseract_cmd = r'/usr/local/Cellar/tesseract/4.1.0/bin/tesseract'


# Global definitions --------------------------------------------------------------------------- ||

class Mode:
    def __init__(self, name, sort_key, scan_area, ship_scan, order_scan, 
                 index_val, shipping_svc, op_select, slips_path, labels_path, scan_area_2 = None):
        self.name         = name
        self.sort_key     = sort_key
        self.scan_area    = scan_area
        self.scan_area_2    = scan_area_2
        self.ship_scan    = ship_scan
        self.order_scan   = order_scan
        self.index_val    = index_val
        self.shipping_svc = shipping_svc
        self.op_select    = op_select
        self.slips_path   = slips_path
        self.labels_path  = labels_path

class OpSelect(enum.Enum):
    OpExport  = 1
    OpReorder = 2
    OpManual  = 3
    OpExit    = 4


manualEntries = dict()



# Function definitions ------------------------------------------------------------------------- ||

##  \fn     Clear
#   \brief  This function simply clears the screen. Does nothing else.
#
def Clear(): 
    if name == 'nt': 
        _ = system('cls') 
    else: 
        _ = system('clear')

##  \fn     GetManualEntries
#   \brief  This function asks for input to fill the values of each key assigned in the
#           dictionary 'manualVals', then returns the dictionary
#
def GetManualEntry(i):
    manualKeys = ['referenceNum', 'full_name', 'addr_line1', 'addr_line2', 'addr_line3', 'addr_line4']
    manualVals = dict.fromkeys(manualKeys, '')
    
    
    # manualVals['referenceNum'] = input(f'\nEnter last 5 digits of reference num on label {i}: ').upper()

    manualVals['full_name'] = input(f'\nEnter full name on label {i}: ').upper()

    ans = input('Enter address? (y/n) ').lower()
    ans = ans if ans != '' else 'n'
    if ans == 'y':
        manualVals['addr_line1'] = input(f'Enter address on label {i}: ').upper()
    ans = input('Enter city/state/zip? (y/n) ').lower()
    ans = ans if ans != '' else 'n'
    if ans == 'y':
        manualVals['addr_line2'] = input(f'Enter city/state/zip on label {i}: ').upper()

    manualVals.update({'referenceNum': 'N/A'})
    return manualVals

##  \fn     Startup
#   \brief  This is the startup function. It displays a welcome menu, allows the user to exit,
#           otherwise accepts arguments for each option.
#   \return <tuple>
#	
def Startup():
    global manualEntries

    slips_path  = ''
    labels_path = ''
    errorMsg    = ''

    s1 = 'Enter file path for Packing Slips: '
    s2 = 'Enter file path for Shipping Labels: '
    s3 = 'Enter indices from previous run: '

    opCodeIncorrect = True
    while (opCodeIncorrect):
        Clear()
        print('[1] Export packing slip data to CSV')
        print('[2] Match & reorder packing slips by shipping labels')
        print('[3] Match & reorder w/manual entry')
        print('[4] Enter to exit')
        if (len(errorMsg) > 0):
            print(errorMsg)
        try:
            op = int(input('\nSelect Option by Number: '))
            if op == 1:
                slips_path = input(s1).strip()
            elif op == 2 or op == 3:
                slips_path = input(s1).strip()
                labels_path = input(s2).strip()
                if op == 3:
                    iList = [s.strip() for s in input(s3).strip('[]').split(',')]
                    iList = [int(i)-1 for i in iList]
                    for i in iList:
                        manualEntries[i] = GetManualEntry(str(i+1))
            elif op == 4:
                print('Exiting script...')
            else:
                errorMsg = 'Incorrect entry... Try again!'
                continue
            opCodeIncorrect = False
        except ValueError:
            errorMsg = 'Wrong data type... Try again!'
    
    return(op, slips_path, labels_path)

    
def pdf_splitter(path, page_order, fname, temp_dir, no_match_keys):
    pdf = PdfFileReader(path)
    for page in range(pdf.getNumPages()):
        pdf_writer = PdfFileWriter()
        pdf_writer.addPage(pdf.getPage(page))
        
        #if mode.op_select == 2:
        #    output_filename = '{}tmp_{}_page_{}.pdf'.format(temp_dir, fname, page_order[page + 1])
        #else:
        if page in no_match_keys:
            output_filename = 'no_match_page_{}.pdf'.format(page)
        else:
            output_filename = '{}tmp_{}_page_{}.pdf'.format(temp_dir, fname, page_order[page])

        with open(output_filename, 'wb') as out:
            pdf_writer.write(out)
            out.close()

def merger(output_path, input_paths):
    pdf_merger = PdfFileMerger()
 
    for path in input_paths:
        pdf_merger.append(path)
 
    with open(output_path, 'wb') as fileobj:
        pdf_merger.write(fileobj)
        
def hasNumbers(string):
    return(any(char.isdigit() for char in string))

def isAddress(s):
    address = re.compile('^[A-Za-z]?(\d+[A-Za-z]?)+\s([A-Za-z0-9]\s?)+')
    return address.match(s)

def format_string(j):
    j = j.upper().replace(',','')
    return(j)

def city_state(data):
    comma = data.find(',')
    city = format_string(data[0:comma])
    t = data[comma:].strip(",").split()
    state = format_string(t[0])
    zip_code = t[1]
    return(city, state, zip_code)

def strip_Addr(q):
    addr_keys = ['full_name', 'first_name', 'last_name', 'company', 'addr_line1', 
                 'addr_line2', 'addr_line3', 'city', 'state', 'zip_code', 'csz']
    addr_book = dict.fromkeys(addr_keys, '')

    companyIndex = 1
    addr1Index   = 1
    addr2Index   = 2
    addr3Index   = 3

    for i, j in enumerate(q):
        if i == 0:
            try:
                t = j.split()
                addr_book['first_name'] = format_string(t[0])
                addr_book['last_name'] = format_string(t[1 if len(t) == 2 else 2])
                addr_book['full_name'] = format_string(j)
            except:
                addr_book['full_name'] = format_string(j)
        if i == 1 and not(isAddress(j)):
        #if i == 1 and hasNumbers(j) == False:
            #addr_book['addr_line1'] = format_string(j)
            addr_book['company'] = format_string(j)
            addr1Index += 1
            addr2Index += 1
            addr3Index += 1
        elif i == addr1Index:
            addr_book['addr_line1'] = format_string(j)
        elif i == addr2Index and i != len(q) - 1:
            addr_book['addr_line2'] = format_string(j)
        elif i == addr3Index and i != len(q) - 1:
            addr_book['addr_line3'] = format_string(j)
        if i == len(q) - 1:
            CSZ = city_state(j)
            addr_book['city'] = CSZ[0]
            addr_book['state'] = CSZ[1]
            addr_book['zip_code'] = CSZ[2]
            addr_book['csz'] = ' '.join(CSZ)
    return(addr_book)


##  \fn     labels_Ripper
#   \brief  This function take the text extracted from a shipping label and identifies
#           the important information: the name, shipping address, and the city/state
#           and zip code. These three lines are return as a list.
#   \return <list>
#
def labels_Ripper(text):
    # Define some variables used
    label_keys = ['full_name', 'addr_line1', 'addr_line2', 'addr_line3', 'addr_line4']
    label_vals = dict.fromkeys(label_keys, '')
    addr = text.splitlines()

    # Setup our regex patterns
    name      = re.compile('([A-Za-z].?\s?)+$')
    address   = re.compile('^[A-Za-z]?(\d+[A-Za-z]?)+\s([A-Za-z0-9]\s?)+')
    cityStZip = re.compile('([A-Za-z]\s?)+,?\s[A-Z]{2}\s\d{5}')

    try:
        # First replace any incorrect characters, these are errors of OCR
        addr = [w.replace('$', 'S') for w in addr]
        addr = [w.replace('§', 'S') for w in addr]

        # Next search for the name
        newAddr = list(filter(name.match, addr))
        # filter out any bad lines we don't want, since name pattern is less restrictive
        newAddr = list(filter(lambda x: 'APTS' not in x, newAddr))
        newAddr = list(filter(lambda x: 'APT' not in x, newAddr))

        # Finally, search for address and city/state/zip
        newAddr.append(list(filter(address.match, addr))[0])
        newAddr.append(list(filter(cityStZip.match, addr))[0])

        if len(newAddr) != 3:
            raise

        # If the lines were successfully found, save them and return
        label_vals['full_name'] = format_string(newAddr[0])
        for i in range(1, len(newAddr)):
            temp = newAddr[i].rstrip()
            temp = temp.replace('-  ', '-' )
            temp = temp.replace('- ', '-' )
            label_vals[label_keys[i]] = temp 
    except Exception as e:
        label_vals['full_name'] = 'Label_Error'

    return(label_vals)


def page_Mod(order_dict):
    # modify numbering format:
    for k, v in order_dict.items():
        if order_dict[k] < 10:
            order_dict[k] = '000' + str(order_dict[k])
        elif order_dict[k] >= 10 and order_dict[k] < 100:
            order_dict[k] = '00' + str(order_dict[k])
        else:
            order_dict[k] = '0' + str(order_dict[k])
    return(order_dict)



def read_reference_number_ups(temp_dir, tmp_name, i):
    # try to get reference number
    tmp_crop_ref = temp_dir + str(i) + '_crop_ref_ups' + '.jpg'
    coords = (0, 2820, 700, 2970)
    Image.open(tmp_name).crop(coords).convert('L').save(tmp_crop_ref)
    # img = cv2.imread(tmp_crop_ref)
    # ret,img = cv2.threshold(img,8,255,cv2.THRESH_BINARY)
    # kernel = np.ones((5,5), np.uint8)
    # img = cv2.morphologyEx(img, cv2.MORPH_OPEN, kernel)
    # img = cv2.bitwise_not(img)
    # img= cv2.dilate(img,kernel,iterations=1)
    # img = cv2.bitwise_not(img)
    # cv2.imwrite(tmp_crop_ref, img)
    
    text = str(((pytesseract.image_to_string(Image.open(tmp_crop_ref)))))
    # print("\nPytesseract: {}".format(text))
    fullRefNoSplit = text.split('Trx Ref No.: ')
    partialRefNoSplit = text.split('No')
    if len(fullRefNoSplit) > 1:
        text = fullRefNoSplit[1]
        text = text.split('\n')[0]
        text = text.replace(' ', '')
        text = text[-6:-1]
        # print("\nPage {} Ref Num: {}".format(i, text))
    elif len(partialRefNoSplit) > 1:
        text = partialRefNoSplit[1]
        text = text.split('\n')[0]
        text = text.replace(' ', '')
        text = text[-6:-1]
        # print("\nPage {} Ref Num: {}".format(i, text))
    else:
        text = "N/A"

    return text

def read_reference_number_fedex(temp_dir, tmp_name, i):
    # try to get reference number
    tmp_crop_ref = temp_dir + str(i) + '_crop_ref_fedex' + '.jpg'
    coords = (792, 842, 1122, 912)
    Image.open(tmp_name).crop(coords).convert('L').save(tmp_crop_ref)
    text = str(((pytesseract.image_to_string(Image.open(tmp_crop_ref)))))
    # print("\nPytesseract: {}".format(text))
    fullRefNoSplit = text.split('REF:')
    
    if len(fullRefNoSplit) > 1:
        text = fullRefNoSplit[1]
        text = text.split('\n')[0]
        text = text.replace(' ', '')
        text = text[-6:-1]
    else:
        text = "N/A"

    # print("\nPage {} Ref Num: {}".format(i, text))

    return text

##  \fn     Read_Labels
#   \brief  This function opens the labels and crops them to the shipping address
#           portion to extract the important information. For certain vendors,
#           different label formats occur, so multiple crop coordinates can be defined.
#   \return <dict>
#
def Read_Labels():
    cropCoordsBedB   = [(64, 350, 1700, 810)]
    cropCoordsBelk   = [(200, 850, 1300, 1210)]

    target1 = (70, 350, 1700, 820) # Fedex Home Delivery
    target2 = (70, 400, 1700, 820) # Fedex Home Delivery
    target3 = (144, 407, 1950, 730)
    cropCoordsTarget = [target1, target2, target3]

    cropCoords = {"bedbath": cropCoordsBedB, "belk": cropCoordsBelk, "target": cropCoordsTarget}

    c_labels = dict()
    label_errors = []

    # remove residual temp dir or create new
    temp_dir = 'labels_temp/'
    if os.path.isdir(temp_dir):
        shutil.rmtree(temp_dir)
    os.mkdir(temp_dir)

    # convert label pdfs to jpgs:
    pages = convert_from_path(mode.labels_path, dpi=500, grayscale=True)
    print('\nProcessing shipping labels...')
    # print(len(pages))
    for i, page in enumerate(pages):
        tmp_name = temp_dir + str(i) + '.jpg'
        tmp_crop = temp_dir + str(i) + '_crop' + '.jpg'

        

        if i in manualEntries.keys():
            c_labels[i] = manualEntries[i]
            # print("\nUsing mangual entry for: page: ", i)
        else:
            
            success = False
            coords_count = 1
            page.save(tmp_name, 'JPEG')

            
            refNum = read_reference_number_ups(temp_dir, tmp_name, i)
            if refNum == "N/A":
                refNum = read_reference_number_fedex(temp_dir, tmp_name, i)
            
            c_labels[i] = { 'pageNum': i+1, 'referenceNum':  refNum }

            for coords in cropCoords.get(mode.name.lower()):
            
                text = str(((pytesseract.image_to_string(Image.open(tmp_name).crop(coords).convert('L')))))

                temp = labels_Ripper(text)
                
                coords_count += 1
                if temp['full_name'] != 'Label_Error':
                    success = True
                    break

                

            c_labels[i].update(temp)
        ref = i + 1
        sys.stdout.write(f'\rReading Shipping Label: %d/{len(pages)}' % ref)
        sys.stdout.flush()
        
    shutil.rmtree(temp_dir)


    # process label error to new file:
    for i, j in c_labels.items():
        if j['full_name'] == 'Label_Error':
            label_errors.append(i)
            
    # if len(label_errors) > 0:
    #     print(f'\n\n >>Label Read Error... Printing {len(label_errors)} labels to Label_Errors.pdf')
    #     print(f'\nIndex values for manual entry are as follows: {[i+1 for i in label_errors]}')
    #     temp_dir = 'lab_errors_temp/'
    #     if os.path.isdir(temp_dir):
    #         shutil.rmtree(temp_dir)
    #     os.mkdir(temp_dir)
        
    #     errors = convert_from_path(mode.labels_path, dpi=200)
    #     for idx, err in enumerate(errors):
    #         if idx in label_errors:
    #             img_path = temp_dir + str(idx) + '.jpg'
    #             err.save(img_path, 'JPEG')
    #             img = Image.open(img_path)
    #             label_errors[label_errors.index(idx)] = img

    #     lab_err_filename = 'Label_Errors.pdf'
    #     label_errors[0].save(lab_err_filename, 'PDF' ,resolution=100.0, save_all=True, append_images=label_errors[1:])
    #     shutil.rmtree(temp_dir)

    print('')
    return(c_labels)

##  \fn     BelkSkuLookup
#   \brief  This function performs a lookup to translate belk vendor SKUs to the corresponding
#           manufacturer's SKU.
#   \params <float64>
#   \return <string>
#
def BelkSkuLookup(vendorSKU):
    vendorCol       = 1
    manufacturerCol = 7

    if skus is None:
        return vendorSKU
    sheet = skus.active

    for row in sheet.iter_rows(min_row=3):
        if row[vendorCol].value == vendorSKU:
            return row[manufacturerCol].value
    
    return ''

def csv_XP(order_info, d, shipping, addr_dict, file_name):

    xport_list = []
    print(f'\nExporting addresses to CSV...')

    for k, v in addr_dict.items():
        if mode.name == 'BedBath':
            mode.shipping_svc = shipping.iloc[k]
            
        xport_load = {
                'Order Number' : order_info.Order_Number.iloc[k],
                'Item Quantity' : int(d[k]['QTY'][0]),
                'Item Marketplace ID' : d[k][mode.sort_key][0],
                'Order Requested Shipping Service' : mode.shipping_svc,
                'Recipient Full Name' : v['full_name'],
                'Recipient First Name' : v['first_name'],
                'Recipient Last Name' : v['last_name'],
                'Recipient Company' : v['company'],
                'Address Line 1' : v['addr_line1'],
                'Address Line 2' : v['addr_line2'],
                'Address Line 3' : v['addr_line3'],
                'City' : v['city'],
                'State' : v['state'],
                'Postal Code' : v['zip_code']
            }
        if mode.name == 'Belk':
            xport_load['Item Marketplace ID'] = BelkSkuLookup(xport_load['Item Marketplace ID'])
        xport_list.append(xport_load)

        if d[k]['Unique_Skews'][0] > 1:
            xplc = xport_load.copy()
            for i in range(1, d[k]['Unique_Skews'][0]):
                xplc['Item Marketplace ID'] = d[k][mode.sort_key][i]
                if mode.name == 'Belk':
                    xplc['Item Marketplace ID'] = BelkSkuLookup(xplc['Item Marketplace ID'])
                xplc['Item Quantity'] = int(d[k]['QTY'][i])
                xport_list.append(xplc.copy())

    df_xport = pd.DataFrame(xport_list, columns = cols)
    df_xport.to_csv(file_name + '_Export.csv')
    print('Export Complete!')

def match_Labels(addr_dict, clean_labels):
    order_dict = dict()
    no_match_keys = []
    dupes_idx = []
    no_match_pages = []
    val_list_names = [k['full_name'] for i, k in addr_dict.items()]
    val_list_addr1 = [k['addr_line1'] for i, k in addr_dict.items()]
    val_list_addr2 = [k['addr_line2'] for i, k in addr_dict.items()]
    val_list_csz = [k['csz'] for i, k in addr_dict.items()]
    val_list_referenceNums = [k['referenceNum'] for i, k in addr_dict.items()]
    
    # catch duplicate names
    for a, b in enumerate(val_list_names):
        if val_list_names.count(b) > 1:
            dupes_idx.append(a)

    for j, k in clean_labels.items():
        if k['referenceNum'] in val_list_referenceNums:
            order_dict[j] = val_list_referenceNums.index(k['referenceNum'])  
        elif k['full_name'] in val_list_names:
            order_dict[j] = val_list_names.index(k['full_name'])  
            dup = order_dict[j]
        elif k['addr_line1'] in val_list_addr1:
            order_dict[j] = val_list_addr1.index(k['addr_line1'])
        elif k['addr_line2'] in val_list_csz:
            order_dict[j] = val_list_csz.index(k['addr_line2'])
        elif k['addr_line3'] in val_list_csz:
            order_dict[j] = val_list_csz.index(k['addr_line3']) 
        else:
            print('No slip for: ', k)
            no_match_pages.append(k['pageNum'])
        try:
            dup = order_dict[j]
            if dup in dupes_idx:   # remove duplicate vals
                val_list_names[dup] = 'xxx'
        except:
            pass
            
    order_keys = list(order_dict.values())
    for k, v in addr_dict.items():
        if k not in order_keys:
            no_match_keys.append(k)
                
    print(f' \nPages to be printed to No_Match_Slips: {len(no_match_keys)}')
    return(order_dict, no_match_keys, no_match_pages)

def bb_ship():
    # collect shipping info: this is for BedBath
    shipping = tabula.read_pdf(mode.slips_path, area = mode.order_scan, pages = 'all', pandas_options={'header': None})
    order_info = pd.DataFrame({'Order_Number':shipping[1].iloc[::3].values}) #review this line
    
    shipping = shipping.iloc[:,-1:]
    shipping.rename(columns= {3:'Shipping'}, inplace = True)
    shipping.drop(shipping[shipping.Shipping.isna()].index, inplace=True)
    shipping.reset_index(drop = True, inplace = True)
    shipping = shipping.Shipping

    # process recipient info:
    ship_to = tabula.read_pdf(mode.slips_path, area = mode.ship_scan, stream = False, pages = 'all', header = None)
    ship_to.rename(columns= {'Shipped To:' : 'Shipping_Addr'}, inplace = True)
    
    #mode.shipping_svc = shipping.iloc[k]
    return(order_info, shipping, ship_to)

def belk_ship():
    order_info = tabula.read_pdf(mode.slips_path, area = mode.order_scan, pages = 'all', pandas_options={'header': None})
    order_info.rename(columns = {1:'Order_Number'}, inplace = True)
    order_info = order_info[::2]

    ship_to = tabula.read_pdf(mode.slips_path, area = mode.ship_scan, pages = 'all', pandas_options={'header': None})
    ship_to.rename(columns= {0: 'Shipping_Addr'}, inplace = True)
    ship_to.drop(ship_to[ship_to.Shipping_Addr.isna()].index, inplace=True)
    ship_to = ship_to[1:].reset_index(drop = True)

    return (order_info, ship_to)

def tar_ship():
    order_info = tabula.read_pdf(mode.slips_path, area = mode.order_scan, pages = 'all', pandas_options={'header': None})
    order_info.rename(columns = {1:'Order_Number'}, inplace = True)
    
    ship_to = tabula.read_pdf(mode.slips_path, area = mode.ship_scan, pages = 'all', pandas_options={'header': None})
    ship_to.rename(columns= {1: 'Shipping_Addr'}, inplace = True)
    ship_to.drop(columns = 0, inplace = True)
    ship_to.drop(ship_to[ship_to.Shipping_Addr.isna()].index, inplace=True)
    ship_to = ship_to[1:].reset_index(drop = True)
    
    return(order_info, ship_to)

def hibbett_ship():
    order_info = tabula.read_pdf(mode.slips_path, area = mode.order_scan, pages = 'all', pandas_options={'header': None})
    order_info.rename(columns = {1:'Order_Number'}, inplace = True)
    
    ship_to = tabula.read_pdf(mode.slips_path, area = mode.ship_scan, pages = 'all', pandas_options={'header': None})
    ship_to.rename(columns= {1: 'Shipping_Addr'}, inplace = True)
    ship_to.drop(columns = 0, inplace = True)
    ship_to.drop(ship_to[ship_to.Shipping_Addr.isna()].index, inplace=True)
    ship_to = ship_to[1:].reset_index(drop = True)
    
    return(order_info, ship_to)

def slips_Reorder(d):
    df = pd.concat(d.values(), ignore_index=True)
    
    df_mul = df[df.Unique_Skews > 1]
    df_mul.drop_duplicates('Page', inplace = True)
    df_mul.Page.astype(int)
    df_mul['Order'] = df_mul.Page / 1000
    
    df_sin = df[df.Unique_Skews == 1]
    df_sin.sort_values(by = mode.sort_key, inplace = True)
    df_sin.reset_index(drop = True, inplace = True)
    df_sin['Order'] = df_sin.index + 1

    df2 = pd.concat([df_mul, df_sin], ignore_index=True)
    df2.Order = df2.index
    
    df2['NewOrder'] = ''
    for idx, row in df2.iterrows():
        if idx < 10:
            temp = '000' + str(row.Order)
        elif (idx >= 10 and idx < 100):
            temp = '00' +  str(row.Order)
        else:
            temp = '0' + str(row.Order)
        df2.at[idx,'NewOrder'] = temp
    return(df2)

def slip_Sort(temp_dir):
    # sort filenames
    slips = os.listdir(temp_dir)
    slips.sort()
    if '.ipynb_checkpoints' in slips:
        slips.remove('.ipynb_checkpoints')
    slips = [temp_dir + str(i) for i in slips]
    return(slips)


def run_reorder(mode, total_pages, addr_dict, file_name, temp_dir):
    
    temp_dir = 'pdf_rip_temp/'
    if os.path.isdir(temp_dir):
        shutil.rmtree(temp_dir)
    os.mkdir(temp_dir)

    clean_labels = Read_Labels()
    
    # create lists from dict keys: 
    name_list = [k['full_name'] for i, k in clean_labels.items()]
    addr1_list = [k['addr_line1'] for i, k in clean_labels.items()]
    addr2_list = [k['addr_line2'] for i, k in clean_labels.items()]
    addr3_list = [k['addr_line3'] for i, k in clean_labels.items()]

    if len(clean_labels) != total_pages:
        print('\nWARNING: # of Shipping Labels does not equal # Packing Slips.')
    
    order_dict, no_match_keys, no_match_pages = match_Labels(addr_dict, clean_labels)
    
    check1 = len(no_match_keys)
    check2 = len(order_dict)

    if check1 + check2 != total_pages:
        print(f'\nREAD ERROR: {check1} packing slips and {check2} labels do not total to {total_pages} input file.\n')
    
    order_list = list(order_dict.values())
    order_dict = {k: v for v,k in enumerate(order_list)}

    # modify numbering format:
    order_dict = page_Mod(order_dict)
    
    # separate packing slips into individual files
    pdf_splitter(mode.slips_path, order_dict, file_name, temp_dir, no_match_keys)

    # sort packing slips
    slips = slip_Sort(temp_dir)

    # merge packing slips to new PDS:
    merger(file_name + '_Reordered.pdf', slips)
    
    # merge matchless files into PDF
    matchless = glob.glob('no_match_page_*.pdf')
    if len(matchless) > 0:
        merger(file_name + '_No_Match_Slips.pdf', matchless)
    
    for file in matchless:
        os.remove(file)
    shutil.rmtree(temp_dir, ignore_errors=True)

    print(file_name + ' Reorder Complete!')
    end_time = time.time()
    print("Time elapsed: {} seconds".format(round(end_time - start_time)))
    return no_match_pages

# Main script section -------------------------------------------------------------------------- ||
start_time = time.time()
bedbath_item_scan = (160.52, 20.12, 240.91, 573.59)

belk_item_scan = (125, 10, 300, 585)
belk_ship_scan = (5, 125, 50, 250)
belk_order_scan = (65, 65, 100, 250)

cols = ['Order Number', 'Order Created Date', 'Order Date Paid', 'Order Total',
   'Order Amount Paid', 'Order Tax Paid', 'Order Shipping Paid',
   'Order Requested Shipping Service', 'Order Total Weight (oz)',
   'Order Custom Field 1', 'Order Custom Field 2', 'Order Custom Field 3',
   'Order Source', 'Order Notes from Buyer', 'Order Notes to Buyer',
   'Order Internal Notes', 'Order Gift Message', 'Order Gift - Flag',
   'Buyer Full Name', 'Buyer First Name', 'Buyer Last Name', 'Buyer Email',
   'Buyer Phone', 'Buyer Username', 'Recipient Full Name',
   'Recipient First Name', 'Recipient Last Name', 'Recipient Phone',
   'Recipient Company', 'Address Line 1', 'Address Line 2',
   'Address Line 3', 'City', 'State', 'Postal Code', 'Country Code',
   'Item SKU', 'Item Name', 'Item Quantity', 'Item Unit Price',
   'Item Weight (oz)', 'Item Options', 'Item Warehouse Location',
   'Item Marketplace ID']

# call startup to define initial variables:
op_select, slips_path, labels_path = Startup()
op_select = OpSelect(op_select)

if op_select == OpSelect.OpExit:
    exit()

addr_index = [0]
d = dict()
addr_dict = dict()
file_name = os.path.splitext(os.path.basename(slips_path))[0]

# automated mode selector
if slips_path.lower().find('target') != -1:
    mode = Mode('Target', 'MFG ID',
                (210, 10, 400, 575), (100, 250, 225, 575), (5, 350, 15, 575),
                'SEND TO:', '', op_select, slips_path, labels_path)

elif slips_path.lower().find('belk') != -1:
    mode = Mode('Belk', 'Item Number', 
                belk_item_scan, belk_ship_scan, belk_order_scan, 
                'Ship To:', '', op_select, slips_path, labels_path)
elif slips_path.lower().find('hibbett') != -1:
    mode = Mode('Hibbett', 'Item Number', 
                (210, 10, 240, 575), (100, 250, 225, 575), (5, 350, 15, 575),
                'Ship To:', '', op_select, slips_path, labels_path, (240, 10, 280, 575))
else:
    mode = Mode('BedBath', 'Vendor Part #', 
                bedbath_item_scan, (600, 305, 750, 600), (10, 200, 75, 600), 
                'Shipped To:', '', op_select, slips_path, labels_path)

print('\nMode:', mode.name)

# get number of packing slips
with open(mode.slips_path, "rb") as filehandle:
    pdf = PdfFileReader(filehandle)
    total_pages = pdf.getNumPages()
    filehandle.close()

print('Processing packing slips...')

# try to get reference number
temp_dir = 'package_slip_temp/'
if os.path.isdir(temp_dir):
    shutil.rmtree(temp_dir)
os.mkdir(temp_dir)


for i in range(total_pages):
    ref = i + 1
    temp = tabula.read_pdf(mode.slips_path, area = mode.scan_area, stream = False, pages = ref)
    if mode.scan_area_2 is not None:
        temp2 = tabula.read_pdf(mode.slips_path, area = mode.scan_area_2, stream = False, pages = ref)
        d[i] = pd.concat([temp,temp2], axis=1)
        d[i].rename(columns= {'PO NUMBER': 'referenceNum'}, inplace = True)
    else:
        d[i] = temp
    # print(d[i])
    d[i].rename(columns= {'Qty': 'QTY'}, inplace = True)
    d[i].rename(columns= {'QTY ORD': 'QTY'}, inplace = True)
    d[i].drop(d[i][d[i].QTY.isna()].index, inplace=True)
    if mode.name == 'Belk':
        d[i].drop(d[i][d[i].QTY.str.isalpha()].index, inplace=True)
    d[i]['Page'] = ref
    d[i]['Unique_Skews'] = d[i].shape[0]

    if 'referenceNum' not in d[i]:
        tmp_name = temp_dir + str(i) + '.jpg'
        pages = convert_from_path(mode.slips_path, dpi=250, grayscale=True)
        page = pages[i]
        page.save(tmp_name, 'JPEG')
        tmp_crop = temp_dir + str(i) + '_crop' + '.jpg'
        coords = (1475, 320, 1875, 400)
        text = str(((pytesseract.image_to_string(Image.open(tmp_name).crop(coords)))))
        releaseNumSplit = text.split('RELEASE # ')
        if len(releaseNumSplit) > 1:
            text = releaseNumSplit[1]
            text = text.split('\n')[0].replace(' ', '')
            text = text[-6:-1]
        else:
            text = "Not Found"
        d[i]['referenceNum'] = text

    d[i].reset_index(inplace = True, drop = True)

    sys.stdout.write(f'\rReading Packing Slip: %d/{total_pages}' % ref)
    sys.stdout.flush()

shutil.rmtree(temp_dir)

if mode.op_select != OpSelect.OpExport:
    # create temp directory / remove residual verions
    temp_dir = 'pdf_rip_temp/'
    if os.path.isdir(temp_dir):
        shutil.rmtree(temp_dir)
    os.mkdir(temp_dir)

# option 2 process:
#if mode.op_select == 2:
#    # use dataframe to clean & reorder labels
#    df2 = slips_Reorder(d)
#    order_dict = dict(zip(df2.Page, df2.NewOrder))
#    # divide packing slips and assign file names corresponding to new order
#    pdf_splitter(mode.slips_path, order_dict, file_name, temp_dir, [])
#    # sort packing slips by filename and combine with merger 
#    slips = slip_Sort(temp_dir)
#    merger(file_name + '_Reordered.pdf', slips)
#    # remove directory containing temporary files 
#    shutil.rmtree(temp_dir)
    
if mode.name == 'BedBath':
    # collect shipping info: this is for BedBath
    order_info, shipping, ship_to = bb_ship()
    #mode.shipping_svc = shipping.iloc[k]
if mode.name == 'Belk':
    order_info, ship_to = belk_ship()
    shipping = ''
    try:
        skus = openpyxl.load_workbook('BELK SKUs.xlsx')
    except FileNotFoundError:
        skus = None
if mode.name == 'Target':
    # collect shipping info: this is for Target
    order_info, ship_to = tar_ship()
    shipping = ''
if mode.name == 'Hibbett':
    order_info, ship_to = hibbett_ship()
    shipping = ''

#process shipping data:
for idx, row in ship_to.iterrows():
    if row.Shipping_Addr == mode.index_val:
        addr_index.append(idx)

for i, j in enumerate(addr_index):
    if j == addr_index[-1]:
        temp = ship_to.iloc[addr_index[i]:, :]
        temp = temp.Shipping_Addr.astype(str)
    else:
        temp = ship_to.iloc[addr_index[i]:addr_index[i + 1], :]
        temp = temp.Shipping_Addr.astype(str)
    if i != 0:
        temp = temp[1:]
    addr_dict[i] = temp
    
for k, v in addr_dict.items():
    q = addr_dict[k]
    addr_dict[k] = strip_Addr(q)
    
# iterate over d and add referenceNum to addr_dict
for index, row in d.items():
    # print(row['referenceNum'].iloc[0])
    addr_dict[index].update({'referenceNum': row['referenceNum'].iloc[0]})

# CSV export option
if mode.op_select == OpSelect.OpExport:
    csv_XP(order_info, d, shipping, addr_dict, file_name)

 
# Packing slip reorder option
if mode.op_select == OpSelect.OpReorder:
    no_match_pages = run_reorder(mode, total_pages, addr_dict, file_name, temp_dir)
    if len(no_match_pages) > 0:
        ans = input('Would you like to reorder with manual entry? (y/n) ').lower()
        ans = ans if ans != '' else 'n'
        if ans == 'y':
            for i in no_match_pages:
                manualEntries[i-1] = GetManualEntry(str(i))
            run_reorder(mode, total_pages, addr_dict, file_name, temp_dir)

elif mode.op_select == OpSelect.OpManual:
    run_reorder(mode, total_pages, addr_dict, file_name, temp_dir)
    

    



