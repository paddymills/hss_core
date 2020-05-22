import pandas

EXPORT = r"C:\Users\PMiller1\Documents\SAP\SAP GUI\export.XLSX"

XL_HEADER = {
    "Material Number": "part",
    "Order quantity (GMEIN)": "qty",
    "WBS Element": "wbs",
    "Occurrence": "shipment",
    "Material description": "desc",
}

# read only HEADER columns
xl = pandas.read_excel(EXPORT)[XL_HEADER.keys()]
xl.rename(columns=XL_HEADER, inplace=True)        # rename HEADER columns

# keep PL*, SHT*, Sheet* desc and D-* WBS items
filter1 = xl[xl.desc.str.startswith(('PL', 'SHT', 'Sheet'))]
filter2 = xl[xl.wbs.str.startswith("D-", na=False)]

# create dataframe from filters
df = pandas.merge(filter1, filter2)
df['job'] = df['part'].str.slice(stop=8)    # add job column

print(df.sort_values(by=['desc']))
