import pandas
import numpy

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
    COST_PER_100LB = 50.75
    COST_COL_STR = 'Cost (at $50.75 per 100lb)'

    def float_as_currency(number):
        return '${:,.2f}'.format(number * COST_PER_100LB * 0.01)

    work_orders = [
        '1170155A-1',
        '1170155B-2',
        '1170155C-3',
        '1170155D-4',
        '1170155E-5',
        '1170155E-6',
        '1170155E-7',
        '1170155E-8',
        '1170155E-9',
        '1170155E-10',
        '1170155E-11',
        '1170155F-12',
        '1170155G-13',
        '1170155R-14',
        '1170155S-15',
        '1170155W-17',
    ]

    df = pandas.DataFrame([[numpy.nan, '$0.00']] * len(work_orders),
                          index=work_orders,
                          columns=['TotalWeight', COST_COL_STR])

    for wo in work_orders:
        sql = get_sql_data(wo + '-%')
        total_weight = 0.0

        for k, v in sql.iterrows():
            part, qty, w = v
            total_weight += qty * w

        df.loc[wo, 'TotalWeight'] = total_weight
        df.loc[wo, COST_COL_STR] = float_as_currency(total_weight)

    print(df)

    total = df.agg('sum')
    print('Total: {}'.format(float_as_currency(total['TotalWeight'])))


def subtotal_df(df, groupby_key, qty_key):
    return df.groupby(groupby_key)[qty_key].sum().reset_index().set_index(groupby_key)


def get_sql_data(workorder='1170155%'):
    conn = engine.connect()
    pip = Table('PIPArchive', sndb_metadata)
    select_columns = [
        pip.c.PartName,
        pip.c.QtyInProcess,
        pip.c.TrueWeight,
    ]
    s = select(select_columns).where(pip.c.TransType == 'SN102').where(
        pip.c.WONumber.like(workorder)).where(~pip.c.PartName.like('1170155%_G%-%'))

    df = pandas.read_sql_query(s, conn)

    return df


if __name__ == "__main__":
    main()
