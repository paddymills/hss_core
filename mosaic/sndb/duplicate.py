
import pyodbc

# :: a function for fixing when a part on a program has no work order attached
def sql_duplicate():
    job = input("Job: ")
    old_part = input("Part to copy data from: ")
    new_part = input("Part to copy data to: ")

    if not all((job, old_part, new_part)): return

    cs = "DRIVER={SQL Server};SERVER=HIIWINBL18;DATABASE=SNDBase91;UID=SNUser;PWD=BestNest1445;"
    conn = pyodbc.connect(cs)
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO PIPArchive (
            ArcDateTime, ProgramName, RepeatID,  SheetName,
            PartName,  WONumber, QtyInProcess, PartLength, PartWidth,
            TrueArea, RectArea, NestedArea, TrueWeight, RectWeight,
            CuttingTime, CuttingLength, PierceQty,
            TrueCost, RectCost, NestedCost, MaterialCost,
            ProcessCost, OutsourceCost, DrawingCost,
            LineItemNumber, DueDate, TransType, TotalCuttingTime,
            MasterPartQty, WOState, RevisionNumber, ArchivePacketID,
            ActualDuration
        )
            SELECT
                ArcDateTime, ProgramName, RepeatID,  SheetName,
                REPLACE(PartName, ?, ?),
                WONumber, QtyInProcess, PartLength, PartWidth,
                TrueArea, RectArea, NestedArea, TrueWeight, RectWeight,
                CuttingTime, CuttingLength, PierceQty,
                TrueCost, RectCost, NestedCost, MaterialCost,
                ProcessCost, OutsourceCost, DrawingCost,
                LineItemNumber, DueDate, TransType, TotalCuttingTime,
                MasterPartQty, WOState, RevisionNumber, ArchivePacketID,
                ActualDuration
            FROM PIPArchive
            WHERE PartName=?
    """, old_part.upper(), new_part.upper(), f"{job.upper()}_{old_part.upper()}")
    cur.commit()
    conn.close()

if __name__ == '__main__':
    sql_duplicate()
