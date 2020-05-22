
number = lambda x: float(x.replace(",", ""))

class Record(object):

    def __eq__(self, other):
        if self.qty == 0 or other.qty == 0:
            return False

        if self.part == other.part and self.wbs == other.wbs:
            return True
        return False

    def matchOtherInPlant(self, other):
        if self.qty == 0 or other.qty == 0:
            return False

        if self.part == other.part and self.plant == other.plant:
            return True
        return False


class CogiRecord(Record):

    def __init__(self, data):
        self.mm = data["Material"]
        self.part = self.mm
        self.name = self.part

        self.loc = data["Storage Location"]
        self.wbs = data["WBS Element"]
        self.qty = number(data["Qty in unit of entry"])
        self.unit = data["Unit of Entry"]
        self.plant = data["Plant"]
        self.order = data["Order"]

        self.error = data["Error Text"]

    def __repr__(self):
        return f"{self.mm} | {self.error}"


class StockRecord(Record):

    def __init__(self, data):
        self.mm = data["Material Number"]
        self.part = self.mm
        self.name = self.part

        self.loc = data["Storage Location"]
        self.wbs = data["Special stock number"]
        self.qty = number(data["Unrestricted"])
        self.unit = data["Base Unit of Measure"]
        self.plant = data["Plant"]

    def __repr__(self):
        return f"{self.mm} | {self.loc} | {self.plant} | {self.wbs}"

class OrderRecord(Record):

    def __init__(self, data):
        self.mm = data["Material Number"]
        self.part = self.mm
        self.name = self.part

        self.wbs = data["WBS Element"]
        self.qty = number(data["Order quantity (GMEIN)"])
        self.plant = data["Plant"]
        self.shipment = data["Occurrence"]

    def __repr__(self):
        return f"{self.mm} | {self.qty} | {self.wbs}"
