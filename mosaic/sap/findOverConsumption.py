import pandas
from sqlalchemy import create_engine, MetaData, Table
from sqlalchemy.sql import select

import tqdm

TABLES_TO_LOAD = ['PIPArchive']

XL_PART_TITLE = 'Material Number'
XL_QTY_TITLE = 'Order quantity (GMEIN)'

engine = create_engine(
    'mssql+pyodbc://SNUser:BestNest1445@hiiwinbl18/SNDBase91?driver=SQL+Server')
sndb_metadata = MetaData()
sndb_metadata.reflect(engine, only=TABLES_TO_LOAD)


def main():
    xl = get_excel_subtotalled()
    sql = get_sql_data()

    new = xl.join(sql).assign(diff=lambda x: x[XL_QTY_TITLE]-x['QtyInProcess'])
    gt = new[new['diff'].gt(0)]

    gt.to_excel(r"C:\Users\PMiller1\Downloads\1170155.xlsx")


def subtotal_df(df, groupby_key, qty_key):
    return df.groupby(groupby_key)[qty_key].sum().reset_index().set_index(groupby_key)


def get_excel_subtotalled():
    path = r"C:\Users\PMiller1\Documents\SAP\SAP GUI\export.xlsx"
    xl_file = pandas.ExcelFile(path)
    df = pandas.read_excel(xl_file, 'Sheet1')

    subtotal = subtotal_df(df, XL_PART_TITLE, XL_QTY_TITLE)

    return subtotal


def get_sql_data():
    conn = engine.connect()
    pip = Table('PIPArchive', sndb_metadata)
    s = select([pip.c.PartName, pip.c.QtyInProcess, pip.c.TrueWeight]).where(pip.c.TransType == 'SN102').where(
        pip.c.PartName.like('1170155%')).where(~pip.c.PartName.like('1170155%_G%-%'))

    df = pandas.read_sql_query(s, conn)
    result = subtotal_df(df, 'PartName', 'QtyInProcess')
    result_without_underscore = result.rename(
        mapper=lambda x: x.replace('_', '-'))

    return result_without_underscore


if __name__ == "__main__":
    main()
