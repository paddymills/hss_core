import re

MARK_PATTERN = r"(?P<job>[0-9]{7}[A-Z])_(?P<type>[A-Z]*)(?P<line>[0-9]+)(?P<girder_group>[A-Z]*)-(?P<flg>[A-Z]*)(?P<plate>[0-9]+)"
MARK_REGEX = re.compile(MARK_PATTERN, re.IGNORECASE)

class Part:

    def __init__(self, **kwargs):
        for k, v in kwargs.items():
            _key = k.lower()
            setattr(self, _key, v)

            mark_items = MARK_REGEX.match(self.mark)
            self.line = mark_items.group('line')
            self.girder_group = mark_items.group('girder_group')
            self.flange = mark_items.group('flg')
            self.plate_number = mark_items.group('plate')

    def __repr__(self):
        return self.mark

    def __hash__(self):
        return hash(self.mark)

    def __eq__(self, other):
        return self.mark == other.mark

    def __ne__(self, other):
        return not(self == other)

    @property
    def lgf(self):
        return '{0:02d}{1}-{2}'.format(self.line, self.girder_group, self.flange)

    @property
    def lgfp(self):
        return '{0}{1:d}'.format(self.lgf, self.plate_number)

    @property
    def next_pl(self):
        return '{0}{1:d}'.format(self.lgf, self.plate_number + 1)

    @property
    def gfp(self):
        return '{0}-{1}{2:d}'.format(self.girder_group, self.flange, self.plate_number)

    def is_next(self, comp, ret=False):  # plate2.is_next(plate1)
        if ret:
            return comp.next_pl == self.lgfp
        elif comp.next_pl == self.lgfp:
            return self
        return None

    def in_line(self, comp):
        if comp.lgf == self.lgf:
            return self
        return None
