import pandas as pd
import os
# import glob
import time
import folder_crawler as fc
import re
import datetime_converter as dc

start = time.time()
# inDir = '/Users/yisli/Documents/landlordlady/Tesla/ad hoc projects/part price breakdown/sample data from Niranjan/'
# rootDir = '/Users/yisli/Documents/landlordlady/Tesla/ad hoc projects/part price breakdown/sample file/'
# rootDir = '//teslamotors.com/US/Finance/Cost IQ/Quotes/M3/Drive Unit\Motor-Transmission'
rootDir = '//teslamotors.com/US/Finance/Cost IQ/Quote IQ/Quotes/MS'
allFile = fc.list_files(rootDir)
outDir = '/Users/yisli/Documents/landlordlady/Tesla/ad hoc projects/part price breakdown/output/'

####################### trying to read all the data from multiple files ###########################
# # extension = ['xlsx']
# os.chdir(inDir)
# # allFile = [i for i in glob.glob(os.path.join(inDir, '*.{}').format(extension))]
# allFile = [i for i in glob.glob(os.path.join(inDir, '*.xlsx'))]
# # allFile = [i for i in glob.glob('*.xlsx')]
# print('files are: ', files)



print('root directory is: ', rootDir)
piece_price_col = ['Cost type', 'Component', 'Detail', 'Quantity in assembly', 'Amount (in std UOM)',
                   'Cost', 'Sub-supplier transport', 'Cycle time (min)', 'Pieces per cycle', 'Machine rate per hr',
                   'Number of operators', 'Labor rate per hr', 'Cost.1']
# print(cols)
outData = pd.DataFrame()
# outData = []
f_skipped = pd.DataFrame()
f_skipped = []
# for f in allFile:
for f in allFile:
    print('\n\n\nreading file: ', f)
    # print(f.rsplit('\\', 1)[1][0])
    # xl = pd.ExcelFile(f)

    ###### check if its an excel file & check if it has the tab "sourcing summary"
    # if f.rsplit('/', 1)[1][0] == '~':
    #### a temporary file has their file name start w "~"
    if f.rsplit('\\', 1)[1][0] == '~':
        print('temporary file; not a real file')
        continue
    elif f.rsplit('.', 1)[1] not in ['xlsx', 'XLSX', 'xls', 'XLS']:
        print('not excel file')
        reason = 'not excel file'
        print(reason)
        # f_skipped.append(pd.DataFrame(f, columns='skipped_file'))
        f_skipped.append({'skipped_file': f, 'reason code': reason})
        continue
    else:
        try:
            if ('Sourcing summary' not in pd.ExcelFile(f).sheet_names):
                # print('sourcing summary cnanot be found; skipping the file')
                reason = 'sourcing summary page cnanot be found; skipping the file'
                print(reason)
                f_skipped.append({'skipped_file': f, 'reason code': reason})
                continue
            elif ('Piece Price' not in pd.ExcelFile(f).sheet_names):
                reason = 'piece price page cnanot be found; skipping the file'
                print(reason)
                f_skipped.append({'skipped_file': f, 'reason code': reason})
                continue
        except:
            reason = 'Workbook encrypted'
            print(reason)
            f_skipped.append({'skipped_file': f, 'reason code': reason})
            continue

    # ####################### grabbing header data ######################
    HeaderDataAll = pd.read_excel(f, sheetname='Sourcing summary', header=1, skiprows=1)

    outFile = os.path.join(outDir, 'result.xlsx')

    ##################### left section data in headers
    # print('date is: ', (HeaderDataAll.ix[HeaderDataAll.iloc[:,0] == 'Date',1].values[0]),'hello')
    # print('length of date data is: ', len(HeaderDataAll.ix[HeaderDataAll.iloc[:,0] == 'Date',1].values[0]))
    print('1st col type is : ', type(str(HeaderDataAll.ix[HeaderDataAll.iloc[:,0] == 'Date',1].values[0])))
    QuoteDate = dc.datetime_converter(str(HeaderDataAll.ix[HeaderDataAll.iloc[:,0] == 'Date',1].values[0]))
    CurrencyHeader = HeaderDataAll.ix[HeaderDataAll.iloc[:,0] == 'Quoted currency',1].values[0]
    Supplier = HeaderDataAll.ix[HeaderDataAll.iloc[:,0] == 'Supplier',1].values[0]
    SupplierContact = HeaderDataAll.ix[HeaderDataAll.iloc[:,0] == 'Supplier contact',1].values[0]
    SupplierPhone = HeaderDataAll.ix[HeaderDataAll.iloc[:,0] == 'Phone',1].values[0]
    SupplierEmail = HeaderDataAll.ix[HeaderDataAll.iloc[:,0] == 'Email',1].values[0]
    PlantLocation = HeaderDataAll.ix[HeaderDataAll.iloc[:,0] == 'Plant location',1].values[0]

    ### since date can come in diff formats, we need to convert them if they are not in standard format
    # QuoteDate = dc.datetime_converter(HeaderDataAll.ix[0, 1])
    # print('Date ', QuoteDate)
    # CurrencyHeader = HeaderDataAll.ix[1, 1]
    # print('Currency on summary page', CurrencyHeader)
    # Supplier = HeaderDataAll.ix[2, 1]
    # print('Supplier ', Supplier)
    # SupplierContact = HeaderDataAll.ix[3, 1]
    # SupplierPhone = HeaderDataAll.ix[4, 1]
    # SupplierEmail = HeaderDataAll.ix[5, 1]
    # PlantLocation = HeaderDataAll.ix[6, 1]
    # print('plant location', PlantLocation)

    ##################### right section data in headers
    print('(second section!!!)')
    print(HeaderDataAll.ix[HeaderDataAll.iloc[:,4] == 'Program duration (yrs)',5].values[0])
    PartName = HeaderDataAll.ix[HeaderDataAll.iloc[:,4] == 'Part name',5].values[0]
    PartNumber = HeaderDataAll.ix[HeaderDataAll.iloc[:,4] == 'Part number',5].values[0]
    Program = HeaderDataAll.ix[HeaderDataAll.iloc[:,4] == 'Program',5].values[0]
    AnnualVolume = HeaderDataAll.ix[HeaderDataAll.iloc[:,4] == 'Annual volume',5].values[0]
    ProgramDuration = HeaderDataAll.ix[HeaderDataAll.iloc[:,4] == 'Program duration (yrs)',5].values[0]
    Incoterms = HeaderDataAll.ix[HeaderDataAll.iloc[:,4] == 'Incoterms',5].values[0]
    PaymentTerms = HeaderDataAll.ix[HeaderDataAll.iloc[:,4] == 'Payment terms',5].values[0]
#     PartName = HeaderDataAll.ix[0,5]
#     print('part name ', PartName)
#     PartNumber = HeaderDataAll.ix[1,5]
#     print('PN ', PartNumber)
#     Program = HeaderDataAll.ix[2,5]
#     print('program', Program)
#     AnnualVolume = HeaderDataAll.ix[3,5]
#     ProgramDuration = HeaderDataAll.ix[4,5]
#     print('program duration is ', ProgramDuration)
#     Incoterms = HeaderDataAll.ix[5,5]
#     PaymentTerms = HeaderDataAll.ix[6,5]

    ##################### markup info
    PurchasedPartMarkup = HeaderDataAll.ix[HeaderDataAll.iloc[:,0] == 'Purchased part markup',1].values[0]
    RawMaterialMarkup = HeaderDataAll.ix[HeaderDataAll.iloc[:,0] == 'Raw material markup',1].values[0]
    MiscOverhead = HeaderDataAll.ix[HeaderDataAll.iloc[:,0] == 'Misc. overhead',1].values[0]
    Scrap = HeaderDataAll.ix[HeaderDataAll.iloc[:,0] == 'Scrap',1].values[0]
    Profit = HeaderDataAll.ix[HeaderDataAll.iloc[:,0] == 'Profit',1].values[0]
    SGA = HeaderDataAll.ix[HeaderDataAll.iloc[:,0] == 'SG&A',1].values[0]
    TotalMarkup = HeaderDataAll.ix[HeaderDataAll.iloc[:,0] == 'Total markup',1].values[0]
    print('total markup is: ', TotalMarkup)


    # print(SupplierContact, "'s email is ", SupplierEmail, " and call him at ", SupplierPhone)

    # Type1Cols = ['Purchased part', 'Raw material', 'Misc. overhead', 'Purchased part markup', 'Raw material markup',
    #              'Scrap', 'Profit', 'SG&A', 'Packaging', 'Freight']
    # Type2Cols = ['Raw material processing', 'Assembly']

    ####################### grabbing detail price data ##########################
    PriceDataValid = pd.read_excel(f, sheetname='Piece Price', header=2, skiprows=1)

    for col in PriceDataValid.columns.values:
        if 'Machine rate' in col:
            ToBeUsed = col
        ##### change "extended cost" to "cost.1" for consistency
        if col == 'Extended cost':
            PriceDataValid.rename(columns={'Extended cost': 'Cost.1'}, inplace=True)
        if col == 'Cost\n w/o pigtail':
            PriceDataValid.rename(columns={'Cost\n w/o pigtail': 'Cost.1'}, inplace=True)
        if col == 'Cost (Painted)':
            PriceDataValid.rename(columns={'Cost (Painted)': 'Cost.1'}, inplace=True)

    CurrencyPrice = (ToBeUsed.split('(')[1]).split('/')[0]
    # print('currency on price page: ', CurrencyPrice)

    ############### replace all (XXX/hr) into "per hr"!!!
    PriceDataValid.columns = [re.sub('\(.+/(.+)\)', r'per \1', s) for s in PriceDataValid.columns]
    print('new columns are: ', PriceDataValid.columns.values)

    # IndexToDelete = []

    PriceDataValid = PriceDataValid[PriceDataValid['Cost.1'].notnull() | PriceDataValid['Cost.1'] > 0]

#     for index, row in PriceDataAll.iterrows():
#         if row['Cost type'] in Type1Cols:
#             if pd.isnull(row.Cost) or row.Cost == 0:
#                 IndexToDelete.append(index)
#         if row['Cost type'] in Type2Cols:
#             if pd.isnull(row['Machine rate ']) or row['Machine rate ($/hr)'] == 0:
#                 IndexToDelete.append(index)
#
#     # print(IndexToDelete)
#     PriceDataValid = PriceDataAll.drop(IndexToDelete)
    PriceDataValid['QuoteDate'] = QuoteDate
    PriceDataValid['CurrencyFromSummary'] = CurrencyHeader
    PriceDataValid['Supplier'] = Supplier
    PriceDataValid['SupplierContact'] = SupplierContact
    PriceDataValid['ContactPhone'] = SupplierPhone
    PriceDataValid['ContactEmail'] = SupplierEmail
    PriceDataValid['PlantLocation'] = PlantLocation

    PriceDataValid['PartName'] = PartName
    PriceDataValid['PartNumber'] = PartNumber
    PriceDataValid['Program'] = Program
    PriceDataValid['AnnualVolume'] = AnnualVolume
    PriceDataValid['ProgramDurationsYrs'] = ProgramDuration
    PriceDataValid['Incoterms'] = Incoterms
    PriceDataValid['PaymentTerms'] = PaymentTerms

    PriceDataValid['PurchasedPartMarkup'] = PurchasedPartMarkup
    PriceDataValid['RawMaterialMarkup'] = RawMaterialMarkup
    PriceDataValid['Misc.Overhead'] = MiscOverhead
    PriceDataValid['Scrap'] = Scrap
    PriceDataValid['Profit'] = Profit
    PriceDataValid['SG&A'] = SGA
    PriceDataValid['TotalMarkup'] = TotalMarkup

    PriceDataValid['File Location'] = f
    PriceDataValid['CurrencyFromPiecePrice'] = CurrencyPrice

    outData = outData.append(PriceDataValid)
    print('time elapsed: ', time.time() - start)

cols = ['QuoteDate', 'CurrencyFromSummary', 'Supplier', 'SupplierContact', 'ContactPhone','ContactEmail',
        'PartName', 'PartNumber', 'Program', 'AnnualVolume', 'ProgramDurationsYrs', 'Incoterms', 'PaymentTerms',
        'PurchasedPartMarkup', 'RawMaterialMarkup', 'Misc.Overhead','Scrap','Profit','SG&A','TotalMarkup',
        'Cost type', 'Component', 'Detail', 'Quantity in assembly', 'Amount (in std UOM)', 'Cost',
        'Sub-supplier transport', 'Cycle time (min)', 'Pieces per cycle', 'Machine rate per hr',
        'Number of operators', 'Labor rate per hr', 'Cost.1',  'CurrencyFromPiecePrice', 'File Location']
print('columns before rearrangement: ', outData.columns.values)
outData = outData[cols]

print(PriceDataValid.columns.values)
# print(len(cols))
# reorder the columns to have supplier and partnumber in the front
# cols = outData.columns.tolist()
# cols = cols[-2:] + cols[:-2]
# outData = outData[cols]
outSkipData = pd.DataFrame(f_skipped)

print('writing file...')
outFile = os.path.join(outDir, 'consolidated data.xlsx')
skipFile = os.path.join(outDir, 'skipped files.xlsx')
outData.to_excel(outFile, index=False)
outSkipData.to_excel(skipFile, index=False)
print('done!')
print('time elapsed: ', time.time() - start)















########## ranamign columns ###############
# print(PriceDataAll.columns.values)
# if 'Machine rate ($/hour)' in PriceDataAll.columns.values:
#     PriceDataAll.rename(columns={'Machine rate ($/hour)': 'Machine rate ($/hr)'}, inplace=True)
# if'Machine rate (€/hr)' in PriceDataAll.columns.values:
#     PriceDataAll.rename(columns={'Machine rate (€/hr)': 'Machine rate ($/hr)'}, inplace=True)
# if'Machine rate (RMB/hr)' in PriceDataAll.columns.values:
#     # print(" column 'Machine rate (RMB/hr)' is in the data table!! ")
#     PriceDataAll.rename(columns={'Machine rate (RMB/hr)': 'Machine rate ($/hr)'}, inplace=True)




# # SampleFileName = '20161120 RUCA Piece Price Tooling and EDT Breakdown.xlsx'
# SampleFileName = 'Tesla_PriceBreakdown-1126105-00-B-20170523.xlsx'
#
# inFile = os.path.join(inDir, SampleFileName)
#
# ####################### grabbing header data ######################
# HeaderDataAll = pd.read_excel(inFile, sheetname='Sourcing summary', header=2, skiprows=1)
#
# outFile = os.path.join(outDir, 'result.xlsx')
# PartNumber = HeaderDataAll.columns.values[5]
# Supplier = HeaderDataAll.ix[0,1]
#
#
#
# ####################### grabbing detail price data ##########################
# PriceDataAll = pd.read_excel(inFile, sheetname='Piece Price', header=2, skiprows=1)
# # print(DataAll)
#
# IndexToDelete = []
#
# Type1Cols = ['Purchased part','Raw material', 'Misc. overhead','Purchased part markup','Raw material markup',
#              'Scrap','Profit','SG&A','Packaging','Freight']
# Type2Cols = ['Raw material processing','Assembly']
#
# for index, row in PriceDataAll.iterrows():
#     if row['Cost type'] in Type1Cols:
#         if pd.isnull(row.Cost) or row.Cost == 0:
#             IndexToDelete.append(index)
#     if row['Cost type'] in Type2Cols:
#         if pd.isnull(row['Machine rate ($/hr)']) or row['Machine rate ($/hr)'] == 0:
#             IndexToDelete.append(index)
#
# # print(IndexToDelete)
# PriceDataValid = PriceDataAll.drop(IndexToDelete)
# PriceDataValid['Supplier'] = Supplier
# PriceDataValid['PartNumber'] = PartNumber
#
# # print(DataValid)
#
# # reorder the columns to have supplier and partnumber in the front
# cols = PriceDataValid.columns.tolist()
# cols = cols[-2:] + cols[:-2]
# PriceDataValid = PriceDataValid[cols]
#
# # writer = pd.ExcelWriter('consolidated data.xlsx')
# outFile = os.path.join(outDir, 'consolidated data.xlsx')
# PriceDataValid.to_excel(outFile, index=False)